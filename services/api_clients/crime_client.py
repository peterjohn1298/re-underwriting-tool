"""FBI UCR Crime Data client using local static CSV.

Data source: Marshall Project / FBI UCR "Offenses Known to Law Enforcement"
Coverage: 68 major US cities, most recent year 2015.
Metric: Violent crimes per 100,000 residents.

Risk thresholds based on 2015 national FBI UCR average (373/100k):
  LOW       < 400   (below national average)
  MODERATE  400-700
  HIGH      700-1100
  VERY HIGH > 1100
"""

import os
import logging
import pandas as pd

from config import DATA_DIR

logger = logging.getLogger(__name__)

CRIME_CSV = os.path.join(DATA_DIR, "crime_data.csv")

# 2015 FBI UCR national averages (per 100k)
NATIONAL_VIOLENT_PER_100K = 373.0
NATIONAL_HOMICIDE_PER_100K = 4.9

# Risk band thresholds (violent crime per 100k)
RISK_BANDS = [
    (400,  "LOW",       "Below national average — relatively safe market"),
    (700,  "MODERATE",  "Near national average — typical urban crime levels"),
    (1100, "HIGH",      "Above average — elevated crime, factor into underwriting"),
    (float("inf"), "VERY HIGH", "Significantly elevated crime — carefully assess tenant quality and insurance"),
]

# City name aliases for fuzzy matching
CITY_ALIASES = {
    "new york": "New York City",
    "nyc": "New York City",
    "new york city": "New York City",
    "la": "Los Angeles",
    "sf": "San Francisco",
    "dc": "Washington",
    "washington dc": "Washington",
    "washington d.c.": "Washington",
    "charlotte": "Charlotte-Mecklenburg",
    "mecklenburg": "Charlotte-Mecklenburg",
    "miami dade": "Miami-Dade",
    "prince georges": "Prince George's",
    "saint louis": "St. Louis",
    "st louis": "St. Louis",
}


class CrimeClient:
    """Lookup city crime rates from local FBI UCR static data."""

    def __init__(self):
        self._df = None
        self._load()

    def _load(self):
        if not os.path.exists(CRIME_CSV):
            logger.warning(f"Crime data CSV not found at {CRIME_CSV}")
            return
        try:
            self._df = pd.read_csv(CRIME_CSV)
            logger.info(f"Crime data loaded: {len(self._df)} cities")
        except Exception as e:
            logger.error(f"Failed to load crime data: {e}")
            self._df = None

    def _normalize(self, name: str) -> str:
        return name.strip().lower().replace("-", " ").replace(".", "")

    def _find_city(self, city: str, state: str = "") -> pd.Series | None:
        if self._df is None:
            return None

        city_clean = self._normalize(city)

        # Check aliases first
        aliased = CITY_ALIASES.get(city_clean)
        if aliased:
            city_clean = self._normalize(aliased)

        # Try exact match on normalized city name
        df = self._df.copy()
        df["_city_norm"] = df["city"].apply(self._normalize)
        matches = df[df["_city_norm"] == city_clean]

        if len(matches) == 1:
            return matches.iloc[0]
        if len(matches) > 1 and state:
            state_matches = matches[matches["state"].str.upper() == state.upper()]
            if len(state_matches) >= 1:
                return state_matches.iloc[0]
            return matches.iloc[0]

        # Try contains match
        contains = df[df["_city_norm"].str.contains(city_clean, na=False)]
        if len(contains) >= 1:
            if state:
                state_m = contains[contains["state"].str.upper() == state.upper()]
                if len(state_m) >= 1:
                    return state_m.iloc[0]
            return contains.iloc[0]

        return None

    def _risk_band(self, violent_per_100k: float) -> tuple[str, str]:
        """Return (risk_level, description) for a given violent crime rate."""
        for threshold, level, desc in RISK_BANDS:
            if violent_per_100k < threshold:
                return level, desc
        return "VERY HIGH", RISK_BANDS[-1][2]

    def get_city_crime(self, city: str, state: str = "") -> dict:
        """Look up crime statistics for a city.

        Returns violent crime rate, risk level, comparison to national average,
        and component crime rates (homicide, robbery, assault).
        """
        row = self._find_city(city, state)

        if row is None:
            return {
                "available": False,
                "city": city,
                "state": state,
                "note": f"No crime data available for {city}, {state}. "
                        f"Coverage: 68 major US cities (FBI UCR 2015).",
                "national_violent_per_100k": NATIONAL_VIOLENT_PER_100K,
                "source": "FBI UCR 2015 (via Marshall Project)",
            }

        violent = row.get("violent_per_100k")
        if pd.isna(violent):
            return {
                "available": False,
                "city": city,
                "state": state,
                "matched_name": row.get("original_name", ""),
                "note": "Crime data for this city is incomplete in the FBI UCR dataset.",
                "national_violent_per_100k": NATIONAL_VIOLENT_PER_100K,
                "source": "FBI UCR 2015 (via Marshall Project)",
            }

        violent = float(violent)
        risk_level, risk_desc = self._risk_band(violent)

        vs_national = ((violent - NATIONAL_VIOLENT_PER_100K) / NATIONAL_VIOLENT_PER_100K * 100)

        return {
            "available": True,
            "city": city,
            "state": state,
            "matched_name": row.get("original_name", ""),
            "violent_per_100k": violent,
            "homicide_per_100k": float(row["homicide_per_100k"]) if pd.notna(row.get("homicide_per_100k")) else None,
            "robbery_per_100k": float(row["robbery_per_100k"]) if pd.notna(row.get("robbery_per_100k")) else None,
            "assault_per_100k": float(row["assault_per_100k"]) if pd.notna(row.get("assault_per_100k")) else None,
            "national_violent_per_100k": NATIONAL_VIOLENT_PER_100K,
            "vs_national_pct": round(vs_national, 1),
            "risk_level": risk_level,
            "risk_description": risk_desc,
            "data_year": int(row.get("data_year", 2015)),
            "source": "FBI UCR 2015 (via Marshall Project)",
            "note": "Based on 2015 FBI UCR data — relative rankings are stable over time.",
        }
