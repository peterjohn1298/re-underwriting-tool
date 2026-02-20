"""Generate AI-written investment memo sections using Claude."""

import logging

logger = logging.getLogger(__name__)

_SECTION_HEADERS = {
    "EXECUTIVE SUMMARY": "executive_summary",
    "MARKET ANALYSIS": "market_analysis",
    "PROPERTY OVERVIEW": "property_overview",
    "FINANCIAL HIGHLIGHTS": "financial_highlights",
    "RISK FACTORS": "risk_factors",
    "INVESTMENT RECOMMENDATION": "recommendation",
}


def _build_memo_context(job: dict) -> str:
    """Build structured context string for Claude to write the memo."""
    deal = job.get("deal")
    pf = job.get("results", {})
    metrics = pf.get("metrics", {})
    reversion = pf.get("reversion", {})
    pro_forma_rows = pf.get("pro_forma", [])
    market = job.get("market_data", {}) or {}
    ml = job.get("ml_valuation", {}) or {}
    mc = job.get("monte_carlo", {}) or {}
    rec = job.get("recommendation", {}) or {}
    rent_pred = job.get("rent_prediction", {}) or {}

    def pct(v):
        return f"{v * 100:.1f}%" if v is not None else "N/A"

    def dollar(v):
        if v is None:
            return "N/A"
        if abs(v) >= 1_000_000:
            return f"${v / 1_000_000:.1f}M"
        return f"${v:,.0f}"

    lines = ["=== DEAL DATA ===", ""]

    if deal:
        lines += [
            f"Property Type: {deal.property_type}",
            f"Address: {deal.address}",
            f"Year Built: {deal.year_built}",
            f"Purchase Price: {dollar(deal.purchase_price)}",
            f"Total Units: {deal.total_units}",
            f"Total SF: {deal.total_sf:,.0f}" if deal.total_sf else "",
            f"In-Place Rent: ${deal.in_place_rent:,.0f}/unit/month",
            f"Market Rent: ${deal.market_rent:,.0f}/unit/month",
            f"Occupancy: {pct(deal.occupancy)}",
            f"Hold Period: {deal.hold_period_years} years",
            f"LTV: {pct(deal.ltv)}  Interest Rate: {pct(deal.interest_rate)}",
            f"Deferred Maintenance: {dollar(deal.deferred_maintenance)}",
            f"Planned CapEx: {dollar(deal.planned_capex)}"
            + (f" ({deal.capex_description})" if deal.capex_description else ""),
        ]

    lines += ["", "=== KEY FINANCIAL METRICS ==="]
    if metrics:
        em = metrics.get("equity_multiple")
        dscr = metrics.get("dscr_yr1")
        lines += [
            f"Levered IRR: {pct(metrics.get('levered_irr'))}",
            f"After-Tax IRR: {pct(metrics.get('after_tax_irr'))}",
            f"Unlevered IRR: {pct(metrics.get('unlevered_irr'))}",
            f"Equity Multiple: {em:.2f}x" if em else "Equity Multiple: N/A",
            f"Cash-on-Cash (Yr 1): {pct(metrics.get('cash_on_cash_yr1'))}",
            f"DSCR (Yr 1): {dscr:.2f}x" if dscr else "DSCR (Yr 1): N/A",
            f"Going-In Cap Rate: {pct(metrics.get('going_in_cap_rate'))}",
            f"Exit Cap Rate: {pct(metrics.get('exit_cap_rate'))}",
            f"Yield on Cost: {pct(metrics.get('yield_on_cost'))}",
            f"Price/Unit: {dollar(metrics.get('price_per_unit'))}",
        ]

    if reversion:
        lines += [
            "", "=== EXIT ANALYSIS ===",
            f"Exit Year: {reversion.get('exit_year', 'N/A')}",
            f"Forward NOI at Exit: {dollar(reversion.get('forward_noi'))}",
            f"Gross Sale Price: {dollar(reversion.get('sale_price'))}",
            f"Net Proceeds (pre-tax): {dollar(reversion.get('net_sale_proceeds'))}",
        ]

    if pro_forma_rows:
        lines += ["", "=== PRO FORMA YEARS 1-3 ==="]
        for row in pro_forma_rows[:3]:
            lines.append(
                f"Year {row.get('year')}: EGI={dollar(row.get('effective_gross_income'))}"
                f"  NOI={dollar(row.get('noi'))}  BTCF={dollar(row.get('btcf'))}"
            )

    # Market data
    demo_s = market.get("demographics", {}).get("structured", {})
    caps = market.get("cap_rates", {})
    crime = market.get("crime", {})
    ws = market.get("walkscore", {})
    hud = market.get("hud_fmr", {})
    rc = market.get("rentcast", {})
    rc_stats = rc.get("market_stats", {}) if rc else {}
    rc_avm = rc.get("rent_estimate", {}) if rc else {}

    lines += ["", "=== MARKET DATA ==="]
    if demo_s:
        pop = demo_s.get("population")
        inc = demo_s.get("median_income")
        rent = demo_s.get("median_rent")
        lines += [
            f"Population: {pop:,}" if pop else "",
            f"Median Household Income: ${inc:,}" if inc else "",
            f"Median Gross Rent: ${rent:,}/mo" if rent else "",
            f"Unemployment Rate: {demo_s.get('unemployment_rate', 'N/A')}%",
            f"Renter-Occupied Units: {demo_s.get('renter_pct', 'N/A')}%",
        ]
    if caps.get("treasury_10yr"):
        lines.append(f"10-Year Treasury: {caps['treasury_10yr']}%")
    if caps.get("average_cap_rate"):
        lines.append(f"Market Cap Rate (FRED-derived): {caps['average_cap_rate']}%")
    if caps.get("mortgage_30yr"):
        lines.append(f"30-Year Mortgage Rate: {caps['mortgage_30yr']}%")
    if crime.get("available"):
        lines.append(
            f"Crime Risk: {crime.get('risk_level')} "
            f"({crime.get('violent_per_100k')} violent crimes/100k, "
            f"{crime.get('vs_national_pct', 0):+.1f}% vs. national avg)"
        )
    if ws.get("available"):
        lines.append(
            f"Walk Score: {ws.get('walk_score')} ({ws.get('walk_label')}) | "
            f"Transit: {ws.get('transit_score')} ({ws.get('transit_label')}) | "
            f"Bike: {ws.get('bike_score')}"
        )
    if hud.get("available"):
        fmr = hud.get("fmr_by_bedroom", {})
        lines.append(
            f"HUD FMR 2026: Studio=${fmr.get('efficiency', 'N/A')}  "
            f"1BR=${fmr.get('one_br', 'N/A')}  2BR=${fmr.get('two_br', 'N/A')}  "
            f"3BR=${fmr.get('three_br', 'N/A')}"
        )
        if hud.get("matched_fmr") and deal and deal.in_place_rent:
            diff = (deal.in_place_rent - hud["matched_fmr"]) / hud["matched_fmr"] * 100
            lines.append(
                f"In-Place Rent vs HUD FMR: {diff:+.1f}% "
                f"({'above' if diff > 0 else 'below'} market)"
            )
    if rc_stats.get("available"):
        lines.append(
            f"RentCast Market: avg=${rc_stats.get('average_rent', 'N/A')}/mo  "
            f"median=${rc_stats.get('median_rent', 'N/A')}/mo"
        )
    if rc_avm.get("available") and rc_avm.get("estimated_rent"):
        lines.append(
            f"RentCast AVM Estimate: ${rc_avm['estimated_rent']:,.0f}/mo "
            f"(range: ${rc_avm.get('rent_range_low', '?'):,}–${rc_avm.get('rent_range_high', '?'):,})"
        )

    if rent_pred and not rent_pred.get("error"):
        lines += [
            "", "=== RENT PREDICTION (ML POLYNOMIAL REGRESSION) ===",
            f"Historical Avg Growth: {rent_pred.get('historical_avg', 'N/A')}%/yr",
            f"Predicted Avg Growth: {rent_pred.get('avg_predicted_growth', 'N/A')}%/yr",
            f"Data Source: {rent_pred.get('data_source', 'FRED CPI Shelter')}",
        ]

    if ml and not ml.get("error"):
        lines += [
            "", "=== ML PROPERTY VALUATION (GradientBoosting) ===",
            f"Assessment: {ml.get('assessment')}",
            f"Predicted Total Value: {dollar(ml.get('predicted_total_value'))}",
            f"Actual Purchase Price: {dollar(deal.purchase_price if deal else None)}",
            f"Premium / Discount: {ml.get('premium_discount_pct', 0):+.1f}%",
        ]

    if mc and not mc.get("error") and mc.get("irr_p50") is not None:
        probs = mc.get("probabilities", {})
        lines += [
            "", "=== MONTE CARLO (1,000 simulations) ===",
            f"P10 IRR: {pct(mc.get('irr_p10'))}  |  "
            f"P50 IRR: {pct(mc.get('irr_p50'))}  |  "
            f"P90 IRR: {pct(mc.get('irr_p90'))}",
            f"Mean IRR: {pct(mc.get('irr_mean'))}  Std Dev: {pct(mc.get('irr_std'))}",
        ]
        for label, val in list(probs.items())[:4]:
            lines.append(f"  {label}: {val:.1f}%")

    if rec:
        lines += [
            "", "=== INTEGRATED RECOMMENDATION ===",
            f"Signal: {rec.get('recommendation', 'N/A')} (Score: {rec.get('score', 'N/A')})",
        ]
        for sig in rec.get("signals", [])[:5]:
            lines.append(f"  {sig}")

    return "\n".join(line for line in lines if line is not None)


def generate_ai_memo(job: dict, openai_api_key: str) -> dict:
    """Call OpenAI GPT-4o to write a 6-section investment memo grounded in deal data.

    Returns a dict with keys: executive_summary, market_analysis,
    property_overview, financial_highlights, risk_factors, recommendation,
    raw (full text), model.
    """
    if not openai_api_key:
        return {"error": "OPENAI_API_KEY not configured"}

    try:
        from openai import OpenAI
        client = OpenAI(api_key=openai_api_key)

        context = _build_memo_context(job)
        rec = job.get("recommendation", {}) or {}
        rec_signal = rec.get("recommendation", "HOLD / CONDITIONAL")

        system_prompt = (
            "You are a senior real estate investment analyst at an institutional fund. "
            "You write precise, data-grounded investment memoranda for acquisition committees. "
            "Every claim must be directly supported by the data provided. "
            "Use specific numbers. Write in professional third-person institutional style. "
            "Do not fabricate facts, market conditions, or data not in the provided context."
        )

        user_prompt = f"""Using ONLY the deal data below, write a complete investment memorandum narrative.
Format each section with its header in ALL CAPS on its own line, followed by the content.

{context}

Write exactly these six sections in order:

EXECUTIVE SUMMARY
(2-3 paragraphs: the property, the investment thesis, headline return metrics, and the recommendation signal)

MARKET ANALYSIS
(2-3 paragraphs: macroeconomic environment, local market fundamentals, rent dynamics, and supply/demand context — cite the specific numbers provided above)

PROPERTY OVERVIEW
(1-2 paragraphs: physical description, current operations, occupancy situation, and value-add opportunity if applicable)

FINANCIAL HIGHLIGHTS
(2 paragraphs: specific return metrics with exact numbers, capital structure summary, and exit analysis)

RISK FACTORS
(List 4-6 specific, quantified risk factors — each as a bullet point starting with •)

INVESTMENT RECOMMENDATION
(1 paragraph: clear {rec_signal} recommendation with specific quantitative rationale)

Rules:
- Use specific numbers from the data (IRR %, DSCR, cap rates, rents, crime stats, etc.)
- Do not invent any figures not in the data
- Each section must follow its header immediately (no blank line between header and content)
- Professional, institutional tone throughout"""

        response = client.chat.completions.create(
            model="gpt-4o",
            max_tokens=2500,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        )

        raw_text = response.choices[0].message.content
        sections = _parse_memo_sections(raw_text)
        sections["raw"] = raw_text
        sections["model"] = "gpt-4o"
        logger.info("AI memo generated successfully")
        return sections

    except Exception as e:
        logger.error(f"AI memo generation failed: {e}")
        return {"error": str(e)}


def _parse_memo_sections(text: str) -> dict:
    """Parse Claude's response into named sections."""
    sections = {}
    current_key = None
    current_lines = []

    for line in text.split("\n"):
        stripped = line.strip().upper()
        matched_key = None
        for header, key in _SECTION_HEADERS.items():
            if stripped == header or stripped.startswith(header + " ") or stripped.startswith(header + ":"):
                matched_key = key
                break

        if matched_key:
            if current_key:
                sections[current_key] = "\n".join(current_lines).strip()
            current_key = matched_key
            current_lines = []
        elif current_key:
            current_lines.append(line)

    if current_key:
        sections[current_key] = "\n".join(current_lines).strip()

    return sections
