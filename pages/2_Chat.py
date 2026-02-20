"""AI Deal Chat — Claude/GPT grounded in analysis results."""

import sys
import os
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent.parent))

st.set_page_config(
    page_title="Deal Chat — RE Underwriting",
    page_icon="💬",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background-color: #0D1117; }
[data-testid="stHeader"] { background-color: #0D1117; }
</style>
""", unsafe_allow_html=True)

# ── Guard ─────────────────────────────────────────────────────────────────────
if "results" not in st.session_state:
    st.warning("No analysis results found. Please run an analysis first.")
    if st.button("← Back to Input Form"):
        st.switch_page("streamlit_app.py")
    st.stop()

R = st.session_state["results"]

st.title("💬 AI Deal Assistant")
st.caption("Ask any question about this deal — grounded in your analysis results. Powered by Google Gemini.")


# ── Build system prompt (mirrors app.py _build_chat_system_prompt) ────────────
def _build_chat_system_prompt() -> str:
    deal = R.get("deal")
    pf = R.get("results", {})
    metrics = pf.get("metrics", {})
    reversion = pf.get("reversion", {})
    pro_forma_rows = pf.get("pro_forma", [])
    market = R.get("market_data", {}) or {}
    ml = R.get("ml_valuation", {}) or {}
    mc = R.get("monte_carlo", {}) or {}
    rec = R.get("recommendation", {}) or {}
    lease = R.get("lease_analysis", {}) or {}

    def pct(v):
        return f"{v*100:.1f}%" if v is not None else "N/A"

    def dollar(v):
        return f"${v:,.0f}" if v is not None else "N/A"

    def fmt(v, d=2):
        return f"{v:.{d}f}" if v is not None else "N/A"

    lines = [
        "You are an expert real estate investment analyst and deal assistant.",
        "You have full access to the underwriting results for the following deal.",
        "Answer all questions using ONLY the data provided below — never fabricate numbers.",
        "Be concise, precise, and quantitative. Flag uncertainty where it exists.",
        "",
        "---------------------------------------",
        "DEAL OVERVIEW",
        "---------------------------------------",
    ]

    if deal:
        lines += [
            f"Property Type: {deal.property_type}",
            f"Address: {deal.address}",
            f"Purchase Price: {dollar(deal.purchase_price)}",
            f"Total Units: {deal.total_units}",
            f"In-Place Rent: {dollar(deal.in_place_rent)}/unit/month",
            f"Occupancy: {pct(deal.occupancy)}",
            f"Hold Period: {deal.hold_period_years} years",
            f"LTV: {pct(deal.ltv)}",
            f"Interest Rate: {pct(deal.interest_rate)}",
            f"Revenue Growth Rate: {pct(deal.revenue_growth_rate)}",
        ]

    lines += [
        "",
        "---------------------------------------",
        "KEY FINANCIAL METRICS",
        "---------------------------------------",
        f"Levered IRR: {pct(metrics.get('levered_irr'))}",
        f"Unlevered IRR: {pct(metrics.get('unlevered_irr'))}",
        f"After-Tax IRR: {pct(metrics.get('after_tax_irr'))}",
        f"Equity Multiple: {fmt(metrics.get('equity_multiple'))}x",
        f"Cash-on-Cash (Yr 1): {pct(metrics.get('cash_on_cash_yr1'))}",
        f"DSCR (Yr 1): {fmt(metrics.get('dscr_yr1'))}",
        f"Yield on Cost: {pct(metrics.get('yield_on_cost'))}",
        f"Going-In Cap Rate: {pct(metrics.get('going_in_cap_rate'))}",
        f"Exit Cap Rate: {pct(metrics.get('exit_cap_rate'))}",
        f"Price per Unit: {dollar(metrics.get('price_per_unit'))}",
    ]

    if reversion:
        lines += [
            "",
            "---------------------------------------",
            "EXIT / REVERSION",
            "---------------------------------------",
            f"Exit Year: {reversion.get('exit_year', 'N/A')}",
            f"Forward NOI at Exit: {dollar(reversion.get('forward_noi'))}",
            f"Exit Sale Price: {dollar(reversion.get('sale_price'))}",
            f"Net Sale Proceeds (after tax): {dollar(reversion.get('net_sale_proceeds_after_tax'))}",
        ]

    if pro_forma_rows:
        lines += [
            "",
            "---------------------------------------",
            "PRO FORMA SUMMARY (first 3 years)",
            "---------------------------------------",
        ]
        for row in pro_forma_rows[:3]:
            yr = row.get("year", "?")
            lines.append(
                f"  Year {yr}: EGI={dollar(row.get('effective_gross_income', 0))}"
                f"  Expenses={dollar(row.get('total_expenses', 0))}"
                f"  NOI={dollar(row.get('noi', 0))}"
                f"  BTCF={dollar(row.get('btcf', 0))}"
            )

    # Market data highlights
    demo = market.get("demographics", {}).get("structured", {})
    caps = market.get("cap_rates", {})
    crime = market.get("crime", {})
    ws = market.get("walkscore", {})
    hud = market.get("hud_fmr", {})
    rc = market.get("rentcast", {})
    rc_stats = rc.get("market_stats", {}) if rc else {}
    rc_avm = rc.get("rent_estimate", {}) if rc else {}
    fz = market.get("flood_zone", {})

    lines += ["", "---------------------------------------", "MARKET DATA", "---------------------------------------"]

    if demo:
        lines += [
            f"Population: {demo.get('population', 'N/A'):,}" if demo.get("population") else "Population: N/A",
            f"Median Household Income: {dollar(demo.get('median_income'))}",
            f"Median Gross Rent: {dollar(demo.get('median_rent'))}/mo",
            f"Unemployment Rate: {fmt(demo.get('unemployment_rate'))}%",
        ]
    if caps.get("average_cap_rate"):
        lines.append(f"Market Cap Rate: {fmt(caps['average_cap_rate'])}%")
    if crime.get("available"):
        lines.append(f"Crime Risk: {crime.get('risk_level')} ({crime.get('vs_national_pct', 0):+.1f}% vs national)")
    if ws.get("available"):
        lines.append(f"Walk Score: {ws.get('walk_score')} | Transit: {ws.get('transit_score')} | Bike: {ws.get('bike_score')}")
    if fz.get("available"):
        lines.append(f"FEMA Flood Zone: {fz.get('flood_zone')} — {fz.get('risk_level')}")
    if hud.get("available"):
        fmr = hud.get("fmr_by_bedroom", {})
        lines.append(f"HUD FMR: Studio=${fmr.get('efficiency', '?')} | 1BR=${fmr.get('one_br', '?')} | 2BR=${fmr.get('two_br', '?')} | 3BR=${fmr.get('three_br', '?')}")
    if rc_avm.get("available") and rc_avm.get("estimated_rent"):
        lines.append(f"RentCast AVM: ${rc_avm['estimated_rent']:,.0f}/mo")

    if ml and ml.get("assessment") and not ml.get("error"):
        lines += [
            "", "---------------------------------------", "ML VALUATION", "---------------------------------------",
            f"Assessment: {ml.get('assessment')}",
            f"Predicted Total Value: {dollar(ml.get('predicted_total_value'))}",
            f"Actual Price/Unit: {dollar(ml.get('actual_price_per_unit'))}",
            f"Premium/Discount: {ml.get('premium_discount_pct', 0):+.1f}%",
        ]

    if mc and mc.get("mean_irr") is not None:
        pcts_mc = mc.get("percentiles", {})
        probs = mc.get("probabilities", {})
        p10 = pcts_mc.get("P10")
        p50 = pcts_mc.get("P50 (Median)")
        p90 = pcts_mc.get("P90")
        lines += [
            "", "---------------------------------------", "MONTE CARLO (1,000 simulations)", "---------------------------------------",
            f"IRR: P10={f'{p10:.1f}%' if p10 is not None else 'N/A'}  "
            f"P50={f'{p50:.1f}%' if p50 is not None else 'N/A'}  "
            f"P90={f'{p90:.1f}%' if p90 is not None else 'N/A'}",
            f"Mean IRR: {mc.get('mean_irr'):.1f}%" if mc.get('mean_irr') is not None else "Mean IRR: N/A",
        ]
        for label, val in probs.items():
            lines.append(f"  {label}: {val:.1f}%")

    if rec:
        lines += [
            "", "---------------------------------------", "INTEGRATED RECOMMENDATION", "---------------------------------------",
            f"Signal: {rec.get('recommendation', 'N/A')}",
            f"Score: {rec.get('score', 'N/A')}",
        ]
        for sig in rec.get("signals", []):
            lines.append(f"  • {sig[0]}: {sig[1]} — {sig[2]}")

    if lease and not lease.get("error"):
        ps = lease.get("portfolio_summary", {})
        if ps:
            lines += [
                "", "---------------------------------------", "LEASE ANALYSIS", "---------------------------------------",
                f"Leases Analyzed: {lease.get('lease_count', 'N/A')}",
                f"Weighted Avg Escalation: {ps.get('weighted_avg_escalation', 'N/A')}%/yr",
            ]

    lines += [
        "",
        "---------------------------------------",
        "INSTRUCTIONS",
        "---------------------------------------",
        "- Answer only from the data above. Do not invent numbers.",
        "- If asked about something not in the data, say so clearly.",
        "- Be concise: 2-4 sentences for most answers.",
        "- For numerical questions, cite the exact figure from the data.",
    ]

    return "\n".join(lines)


# ── API key resolution ────────────────────────────────────────────────────────
def _get_google_key() -> str:
    try:
        k = st.secrets.get("GOOGLE_API_KEY", "")
    except Exception:
        k = ""
    return k or os.environ.get("GOOGLE_API_KEY", "")


gemini_key = _get_google_key()
if not gemini_key:
    st.warning("No Google API key configured. Add GOOGLE_API_KEY to your Streamlit secrets.")


# ── Chat UI ───────────────────────────────────────────────────────────────────
if "chat_history" not in st.session_state:
    st.session_state["chat_history"] = []

# Display existing messages
for msg in st.session_state["chat_history"]:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Chat input
prompt = st.chat_input("Ask anything about this deal... (e.g. 'What is the main risk?')")
if prompt and gemini_key:
    # Show user message
    st.session_state["chat_history"].append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # Build conversation history (last 10 turns)
    # Gemini uses role "model" instead of "assistant", and "parts" instead of "content"
    system_prompt = _build_chat_system_prompt()
    gemini_history = []
    for h in st.session_state["chat_history"][-10:]:
        if h["role"] in ("user", "assistant"):
            role = "model" if h["role"] == "assistant" else "user"
            gemini_history.append({"role": role, "parts": [h["content"]]})

    # Call Gemini — streaming so tokens appear word-by-word
    with st.chat_message("assistant"):
        try:
            import google.generativeai as genai
            genai.configure(api_key=gemini_key)
            model = genai.GenerativeModel(
                model_name="gemini-2.0-flash",
                system_instruction=system_prompt,
            )
            response = model.generate_content(gemini_history, stream=True)
            reply = st.write_stream(
                chunk.text for chunk in response if hasattr(chunk, "text") and chunk.text
            )
            st.session_state["chat_history"].append({"role": "assistant", "content": reply})
        except Exception as e:
            err_msg = f"AI error: {str(e)}"
            st.error(err_msg)
            st.session_state["chat_history"].append({"role": "assistant", "content": err_msg})

elif prompt and not gemini_key:
    st.error("Google API key required for chat. Please configure GOOGLE_API_KEY in Streamlit secrets.")

# Clear history button
if st.session_state["chat_history"]:
    if st.button("🗑️ Clear Chat"):
        st.session_state["chat_history"] = []
        st.rerun()
