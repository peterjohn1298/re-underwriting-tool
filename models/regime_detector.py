"""Housing Market Regime Detector.

Classifies the current macro housing environment into one of four regimes
using live FRED data already fetched during market research:

  EXPANSION   — Rising starts, tightening vacancy, elevated rent inflation
  PEAK        — Starts plateauing, vacancy at lows, rent inflation cooling
  CONTRACTION — Starts declining, vacancy rising, rent deflation pressure
  TROUGH      — Starts depressed, high vacancy, rent stagnant / recovering

Each regime produces an investment signal and adjusted risk commentary for
the underwriting recommendation engine.
"""

import logging

logger = logging.getLogger(__name__)


# ── Regime definitions ────────────────────────────────────────────────────────

REGIMES = {
    "EXPANSION": {
        "label":       "EXPANSION",
        "color":       "#3fb950",   # green
        "signal":      "POSITIVE",
        "description": (
            "Housing market is in an expansion phase: starts rising, vacancy "
            "tightening, and rent inflation elevated. Landlord-favorable conditions "
            "support above-trend rent growth assumptions."
        ),
        "underwriting_implication": (
            "Consider modeling rent growth at the upper end of your range. "
            "Exit cap rate compression is plausible. Strong demand supports occupancy."
        ),
    },
    "PEAK": {
        "label":       "PEAK",
        "color":       "#e3b341",   # gold/amber
        "signal":      "NEUTRAL",
        "description": (
            "Housing market appears to be peaking: starts stabilizing, vacancy near "
            "cycle lows, rent inflation beginning to moderate. Late-cycle dynamics "
            "warrant conservative exit cap assumptions."
        ),
        "underwriting_implication": (
            "Use base-case rent growth; stress-test with a Bear scenario. "
            "Exit cap rate expansion likely over a 7–10 year hold."
        ),
    },
    "CONTRACTION": {
        "label":       "CONTRACTION",
        "color":       "#f85149",   # red
        "signal":      "WARNING",
        "description": (
            "Housing market is contracting: starts falling, vacancy rising, rent "
            "inflation cooling or negative. Challenging conditions for rent growth "
            "and occupancy assumptions."
        ),
        "underwriting_implication": (
            "Apply a haircut to rent growth assumptions. Stress-test DSCR at "
            "higher vacancy. Exit cap rate expansion should be the base case."
        ),
    },
    "TROUGH": {
        "label":       "TROUGH",
        "color":       "#58a6ff",   # blue (recovery signal)
        "signal":      "NEUTRAL",
        "description": (
            "Housing market is near a trough: starts depressed, vacancy elevated, "
            "but signs of stabilization emerging. Contrarian opportunity for "
            "long-hold investors with strong balance sheets."
        ),
        "underwriting_implication": (
            "Conservative near-term assumptions are warranted, but model recovery "
            "in years 3–5. Long hold periods may capture mean reversion upside."
        ),
    },
}


# ── Scoring helpers ───────────────────────────────────────────────────────────

def _score_housing_starts(starts_data: dict) -> int:
    """Score -1 to +1 based on housing starts level and trend.

    HOUST is in thousands of units (SAAR). Long-run average ~1,400k.
    Recent trend: compare latest to 6-month-ago reading.
    """
    current = starts_data.get("current")
    if current is None:
        return 0

    recent = starts_data.get("recent_data", [])

    # Level signal
    level_score = 0
    if current >= 1_500:
        level_score = 1      # Above long-run average → expansionary
    elif current <= 1_000:
        level_score = -1     # Well below average → contractionary

    # Trend signal (compare to 6 months ago)
    trend_score = 0
    if len(recent) >= 6:
        older = recent[-1].get("value", current) if recent else current
        change_pct = (current - older) / older * 100 if older else 0
        if change_pct > 5:
            trend_score = 1
        elif change_pct < -5:
            trend_score = -1

    return max(-1, min(1, level_score + trend_score))


def _score_vacancy_rate(vacancy_data: dict) -> int:
    """Score -1 to +1 based on rental vacancy rate.

    Long-run average ~6.5%. Lower vacancy = tighter market = positive.
    """
    current = vacancy_data.get("current_rate")
    if current is None:
        return 0

    if current <= 5.0:
        return 1     # Very tight — expansionary
    elif current <= 6.5:
        return 0     # Near average
    elif current <= 8.0:
        return -1    # Elevated — contractionary
    else:
        return -1    # High — deep contraction/trough


def _score_shelter_cpi(cap_rates_data: dict, market_data: dict) -> int:
    """Score -1 to +1 based on CPI Shelter year-over-year growth.

    Shelter CPI > 5% → hot market (expansionary)
    Shelter CPI 2–5% → moderate (neutral)
    Shelter CPI < 2% → cooling (contractionary)
    """
    # Try to get from market_data rent_trends or cap_rates
    rent_trends = market_data.get("rent_trends", {}) if market_data else {}
    shelter_growth = rent_trends.get("cpi_shelter_yoy")

    if shelter_growth is None:
        # Fallback: infer from cap_rates macro data
        return 0

    if shelter_growth >= 5.0:
        return 1
    elif shelter_growth >= 2.0:
        return 0
    else:
        return -1


def _score_yield_curve(cap_rates_data: dict) -> int:
    """Score based on 10yr-2yr Treasury spread (yield curve).

    Steep curve (>100bps) → growth expectations → positive
    Flat/inverted (<0bps) → recession risk → negative
    """
    t10 = cap_rates_data.get("treasury_10yr")
    t2  = cap_rates_data.get("treasury_2yr")
    if t10 is None or t2 is None:
        return 0

    spread_bps = (t10 - t2) * 100

    if spread_bps > 100:
        return 1
    elif spread_bps > 0:
        return 0
    else:
        return -1    # Inverted — recession signal


def _score_unemployment(demographics: dict) -> int:
    """Score based on unemployment rate.

    Below 4% → tight labor market → expansionary
    4–6% → moderate
    Above 6% → weakening → contractionary
    """
    structured = demographics.get("structured", {}) if demographics else {}
    unemp = structured.get("unemployment_rate")
    if unemp is None:
        return 0

    if unemp <= 4.0:
        return 1
    elif unemp <= 6.0:
        return 0
    else:
        return -1


# ── Main detector ─────────────────────────────────────────────────────────────

class HousingRegimeDetector:
    """Detect the current housing market cycle regime from FRED macro data."""

    def detect(self, market_data: dict) -> dict:
        """Run regime detection from the market_data dict already fetched.

        Args:
            market_data: Output of services/market_research.run_full_research()

        Returns:
            dict with regime label, score, signal, description, implication,
            and per-indicator scores for display.
        """
        if not market_data:
            return {"regime": "UNKNOWN", "error": "No market data available"}

        # market_data["macro"] holds FRED snapshot: housing_starts, rental_vacancy_rate,
        # treasury_10yr, treasury_2yr, cpi_yoy_inflation
        macro     = market_data.get("macro", {}) or {}
        cap_rates = market_data.get("cap_rates", {}) or {}
        demo      = market_data.get("demographics", {}) or {}

        # Build lightweight dicts that the scoring helpers expect
        starts_raw  = {"current": macro.get("housing_starts")}
        vacancy_raw = {"current_rate": macro.get("rental_vacancy_rate")}

        # Merge treasury data — cap_rates may also carry it
        t10 = macro.get("treasury_10yr") or cap_rates.get("treasury_10yr")
        t2  = macro.get("treasury_2yr")  or cap_rates.get("treasury_2yr")
        cap_rates_merged = dict(cap_rates)
        cap_rates_merged["treasury_10yr"] = t10
        cap_rates_merged["treasury_2yr"]  = t2

        # Score each dimension (-1, 0, +1)
        scores = {
            "housing_starts":  _score_housing_starts(starts_raw),
            "vacancy_rate":    _score_vacancy_rate(vacancy_raw),
            "shelter_cpi":     _score_shelter_cpi(cap_rates_merged, market_data),
            "yield_curve":     _score_yield_curve(cap_rates_merged),
            "unemployment":    _score_unemployment(demo),
        }

        total = sum(scores.values())

        # Map composite score to regime  (range: -5 to +5)
        if total >= 2:
            regime_key = "EXPANSION"
        elif total >= 0:
            regime_key = "PEAK"
        elif total >= -2:
            regime_key = "CONTRACTION"
        else:
            regime_key = "TROUGH"

        regime = REGIMES[regime_key].copy()
        regime["composite_score"] = total
        regime["indicator_scores"] = scores
        regime["raw"] = {
            "housing_starts_current": macro.get("housing_starts"),
            "vacancy_rate":           macro.get("rental_vacancy_rate"),
            "treasury_10yr":          t10,
            "treasury_2yr":           t2,
        }

        logger.info(f"Housing regime detected: {regime_key} (score={total})")
        return regime
