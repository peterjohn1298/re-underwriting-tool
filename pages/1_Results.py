"""Results dashboard — all 11 analysis sections."""

import sys
import copy
import os
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent.parent))

st.set_page_config(
    page_title="Results — RE Underwriting",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background-color: #0D1117; }
[data-testid="stHeader"] { background-color: #0D1117; }
.stTabs [data-baseweb="tab-list"] { gap: 4px; background: #0D1117; }
.stTabs [data-baseweb="tab"] {
    background: #161B22; border-radius: 6px;
    padding: 7px 14px; color: #8B949E; font-size: 0.85em;
}
.stTabs [aria-selected="true"] {
    background-color: #C9A84C !important;
    color: #0D1117 !important; font-weight: bold;
}
.verdict-badge {
    display: inline-block; padding: 10px 28px;
    border-radius: 24px; font-size: 1.6em; font-weight: 900; letter-spacing: 2px;
}
.verdict-STRONG-BUY  { background: #1a7f37; color: #fff; }
.verdict-BUY         { background: #2ea043; color: #fff; }
.verdict-HOLD        { background: #9e6a03; color: #fff; }
.verdict-PASS        { background: #da3633; color: #fff; }
.signal-STRONG-BUY, .signal-BUY, .signal-POSITIVE { color: #3fb950; }
.signal-HOLD, .signal-NEUTRAL, .signal-CAUTION     { color: #e3b341; }
.signal-PASS, .signal-WARNING                       { color: #f85149; }
.metric-chip {
    background: #161B22; border: 1px solid #30363D;
    border-radius: 8px; padding: 12px 16px; text-align: center;
}
.chip-label { font-size: 0.78em; color: #8B949E; margin-bottom: 4px; }
.chip-value { font-size: 1.35em; font-weight: 700; color: #E6EDF3; }
</style>
""", unsafe_allow_html=True)

# ── Guard ─────────────────────────────────────────────────────────────────────
if "results" not in st.session_state:
    st.warning("No analysis results found. Please run an analysis first.")
    if st.button("← Back to Input Form"):
        st.switch_page("streamlit_app.py")
    st.stop()

R = st.session_state["results"]
pro_forma  = R.get("results", {})
market     = R.get("market_data", {}) or {}
ml         = R.get("ml_valuation") or {}
mc         = R.get("monte_carlo") or {}
rec        = R.get("recommendation") or {}
lease      = R.get("lease_analysis") or {}
rent_pred  = R.get("rent_prediction") or {}
backtest   = R.get("backtest") or {}
scenarios  = R.get("scenarios") or {}
waterfall  = R.get("waterfall") or {}
ai_memo    = R.get("ai_memo") or {}
unit_mix   = R.get("unit_mix") or {}
sensitivity = R.get("sensitivity") or {}
deal_obj   = R.get("deal")
excel_path = R.get("excel_path")
word_path  = R.get("word_path")
pdf_path   = R.get("pdf_path")

metrics   = pro_forma.get("metrics", {})
reversion = pro_forma.get("reversion", {})
pf_rows   = pro_forma.get("pro_forma", [])
inputs_d  = pro_forma.get("inputs", {})
derived_d = inputs_d.get("derived", {})

def _pct(v):
    return f"{v*100:.1f}%" if v is not None else "—"

def _dollar(v):
    return f"${v:,.0f}" if v is not None else "—"

def _x(v, dec=2):
    return f"{v:.{dec}f}x" if v is not None else "—"

def _fmt(v, dec=2):
    return f"{v:.{dec}f}" if v is not None else "—"


# ── Header ────────────────────────────────────────────────────────────────────
prop_name = derived_d.get("property_name", "Subject Property")
prop_addr = deal_obj.address if deal_obj else ""
st.title(f"📊 {prop_name}")
if prop_addr:
    st.caption(prop_addr)

col_hdr1, col_hdr2 = st.columns([3, 1])
with col_hdr2:
    if st.button("← New Analysis", use_container_width=True):
        st.switch_page("streamlit_app.py")
    # Save for comparison
    if "saved_deals" not in st.session_state:
        st.session_state["saved_deals"] = {}
    job_id = R.get("job_id", "")
    if job_id and job_id not in st.session_state["saved_deals"]:
        if st.button("📌 Save for Comparison", use_container_width=True):
            st.session_state["saved_deals"][job_id] = {
                "job_id": job_id,
                "property_name": derived_d.get("property_name", ""),
                "address": deal_obj.address if deal_obj else "",
                "purchase_price": deal_obj.purchase_price if deal_obj else 0,
                "total_units": deal_obj.total_units if deal_obj else 0,
                "levered_irr": metrics.get("levered_irr"),
                "equity_multiple": metrics.get("equity_multiple"),
                "cash_on_cash_yr1": metrics.get("cash_on_cash_yr1"),
                "dscr_yr1": metrics.get("dscr_yr1"),
                "going_in_cap_rate": metrics.get("going_in_cap_rate"),
                "exit_cap_rate": metrics.get("exit_cap_rate"),
                "annual_nois": pro_forma.get("annual_nois", []),
            }
            st.success("Saved!")

st.markdown("---")

# ── What-If Sidebar ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("🔧 What-If Calculator")
    st.caption("Adjust assumptions to see real-time impact on returns.")

    wi_rent_growth = st.slider("Rent Growth (%/yr)", 0.0, 10.0,
                                float((deal_obj.revenue_growth_rate * 100) if deal_obj else 3.0),
                                step=0.25)
    wi_occupancy = st.slider("Occupancy (%)", 50, 100,
                              int((deal_obj.occupancy * 100) if deal_obj else 92))
    wi_exit_cap = st.slider("Exit Cap Spread (bps)", -100, 200,
                             int((deal_obj.exit_cap_rate_spread * 10000) if deal_obj else 25),
                             step=5)
    wi_expense_growth = st.slider("Expense Growth (%/yr)", 0.0, 10.0,
                                   float((deal_obj.expense_growth_rate * 100) if deal_obj else 3.0),
                                   step=0.25)

    wi_metrics = metrics.copy()
    if deal_obj:
        try:
            from models.financial_model import build_pro_forma
            sim_deal = copy.deepcopy(deal_obj)
            sim_deal.revenue_growth_rate = wi_rent_growth / 100
            sim_deal.occupancy = wi_occupancy / 100
            sim_deal.exit_cap_rate_spread = wi_exit_cap / 10000
            sim_deal.expense_growth_rate = wi_expense_growth / 100
            sim_deal.yearly_revenue_growth = []  # Use flat rate
            wi_result = build_pro_forma(sim_deal)
            wi_metrics = wi_result.get("metrics", metrics)
        except Exception:
            pass

    st.markdown("---")
    st.markdown("**What-If Results**")
    wi_irr = wi_metrics.get("levered_irr")
    wi_em  = wi_metrics.get("equity_multiple")
    wi_coc = wi_metrics.get("cash_on_cash_yr1")
    wi_dscr = wi_metrics.get("dscr_yr1")
    col_w1, col_w2 = st.columns(2)
    col_w1.metric("IRR", _pct(wi_irr))
    col_w2.metric("EM", _x(wi_em))
    col_w1.metric("CoC", _pct(wi_coc))
    col_w2.metric("DSCR", _fmt(wi_dscr))


# ── Main Tabs ─────────────────────────────────────────────────────────────────
tabs = st.tabs([
    "🏆 Overview",
    "💰 Returns",
    "📈 Pro Forma",
    "🌍 Market Data",
    "🤖 ML & Monte Carlo",
    "📉 Rent Forecast",
    "📊 Scenarios",
    "📄 Lease Analysis",
    "🤖 AI Memo",
    "🌊 Waterfall",
    "⬇️ Downloads",
])

# ── Tab 0: Overview ───────────────────────────────────────────────────────────
with tabs[0]:
    # Verdict badge
    verdict = rec.get("recommendation", "HOLD / CONDITIONAL")
    badge_class = "verdict-" + verdict.replace(" / ", "-").replace(" ", "-")
    st.markdown(
        f'<div style="text-align:center;margin:16px 0">'
        f'<span class="verdict-badge {badge_class}">{verdict}</span></div>',
        unsafe_allow_html=True,
    )

    score = rec.get("score", 0)
    st.markdown(f"<div style='text-align:center;color:#8B949E;margin-bottom:20px'>"
                f"Signal Score: {score:+d}</div>", unsafe_allow_html=True)

    # Scorecard chips
    chip_cols = st.columns(5)
    chips = [
        ("Levered IRR",      _pct(metrics.get("levered_irr"))),
        ("Equity Multiple",  _x(metrics.get("equity_multiple"))),
        ("Cash-on-Cash Y1",  _pct(metrics.get("cash_on_cash_yr1"))),
        ("DSCR Y1",          _fmt(metrics.get("dscr_yr1"))),
        ("Going-In Cap",     _pct(metrics.get("going_in_cap_rate"))),
    ]
    for col, (label, value) in zip(chip_cols, chips):
        col.markdown(
            f'<div class="metric-chip">'
            f'<div class="chip-label">{label}</div>'
            f'<div class="chip-value">{value}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # Signal table
    st.subheader("Signal Breakdown")
    signals = rec.get("signals", [])
    if signals:
        for sig in signals:
            source, verdict_s, detail = sig[0], sig[1], sig[2]
            css = f"signal-{verdict_s.replace(' ', '-')}"
            st.markdown(
                f"**{source}** — "
                f'<span class="{css}">{verdict_s}</span> — {detail}',
                unsafe_allow_html=True,
            )
    else:
        st.info("No signals generated (run with optional features enabled for richer analysis).")

    # Snapshot chips row 2
    st.markdown("---")
    snap_cols = st.columns(5)
    snaps = [
        ("Price / Unit",   _dollar(metrics.get("price_per_unit"))),
        ("Yield on Cost",  _pct(metrics.get("yield_on_cost"))),
        ("Unlevered IRR",  _pct(metrics.get("unlevered_irr"))),
        ("After-Tax IRR",  _pct(metrics.get("after_tax_irr"))),
        ("Exit Cap Rate",  _pct(metrics.get("exit_cap_rate"))),
    ]
    for col, (label, value) in zip(snap_cols, snaps):
        col.markdown(
            f'<div class="metric-chip">'
            f'<div class="chip-label">{label}</div>'
            f'<div class="chip-value">{value}</div></div>',
            unsafe_allow_html=True,
        )

    # Unit mix summary (if available)
    if unit_mix and unit_mix.get("units"):
        st.markdown("---")
        st.subheader("Unit Mix")
        import pandas as pd
        um_rows = []
        for u in unit_mix["units"]:
            um_rows.append({
                "Type": u["type_name"],
                "Units": u["count"],
                "In-Place Rent": f"${u['in_place_rent']:,.0f}",
                "Market Rent": f"${u['market_rent']:,.0f}" if u.get("market_rent") else "—",
                "Annual Revenue": f"${u['annual_revenue']:,.0f}",
                "HUD FMR": f"${u['hud_fmr']:,.0f}" if u.get("hud_fmr") else "N/A",
                "vs HUD": f"{u['vs_hud_pct']:+.1f}%" if u.get("vs_hud_pct") is not None else "—",
            })
        st.dataframe(pd.DataFrame(um_rows), use_container_width=True, hide_index=True)
        blended = unit_mix.get("blended_in_place_rent", 0)
        premium = unit_mix.get("total_rent_premium_potential", 0)
        st.caption(
            f"Blended in-place rent: **${blended:,.0f}/unit/mo** | "
            f"Total annual rent-to-market upside: **${premium:,.0f}**"
        )


# ── Tab 1: Returns ────────────────────────────────────────────────────────────
with tabs[1]:
    st.subheader("Return Metrics")
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Levered IRR",    _pct(metrics.get("levered_irr")))
    r2.metric("Unlevered IRR",  _pct(metrics.get("unlevered_irr")))
    r3.metric("After-Tax IRR",  _pct(metrics.get("after_tax_irr")))
    r4.metric("Equity Multiple", _x(metrics.get("equity_multiple")))

    r5, r6, r7, r8 = st.columns(4)
    r5.metric("Cash-on-Cash Y1", _pct(metrics.get("cash_on_cash_yr1")))
    r6.metric("DSCR Y1",         _fmt(metrics.get("dscr_yr1")))
    r7.metric("Yield on Cost",   _pct(metrics.get("yield_on_cost")))
    r8.metric("Stabilized YOC",  _pct(metrics.get("stabilized_yoc")))

    st.markdown("---")
    st.subheader("Capital Structure")
    cs1, cs2, cs3, cs4 = st.columns(4)
    cs1.metric("Purchase Price",    _dollar(deal_obj.purchase_price if deal_obj else None))
    cs2.metric("Loan Amount",       _dollar(derived_d.get("loan_amount")))
    cs3.metric("Equity Required",   _dollar(derived_d.get("equity_required")))
    cs4.metric("Total Project Cost", _dollar(derived_d.get("total_project_cost")))

    st.markdown("---")
    st.subheader("Exit / Reversion")
    ex1, ex2, ex3, ex4 = st.columns(4)
    ex1.metric("Exit Year",         str(reversion.get("exit_year", "—")))
    ex2.metric("Exit Cap Rate",     _pct(reversion.get("exit_cap_rate")))
    ex3.metric("Exit Sale Price",   _dollar(reversion.get("sale_price")))
    ex4.metric("Net Proceeds (AT)", _dollar(reversion.get("net_sale_proceeds_after_tax")))

    # Variable growth indicator
    if rec.get("used_variable_growth"):
        st.info("📈 Pro forma used ML-predicted year-by-year rent growth rates.")


# ── Tab 2: Pro Forma ──────────────────────────────────────────────────────────
with tabs[2]:
    import pandas as pd
    st.subheader("10-Year Pro Forma")
    if pf_rows:
        cols_to_show = [
            "year", "effective_gross_income", "total_expenses",
            "noi", "debt_service", "btcf", "atcf",
        ]
        rename = {
            "year": "Year",
            "effective_gross_income": "EGI ($)",
            "total_expenses": "Expenses ($)",
            "noi": "NOI ($)",
            "debt_service": "Debt Service ($)",
            "btcf": "BTCF ($)",
            "atcf": "ATCF ($)",
        }
        rows = []
        for row in pf_rows:
            r = {k: row.get(k, 0) for k in cols_to_show if k in row}
            rows.append(r)
        df = pd.DataFrame(rows).rename(columns=rename)
        for col in df.columns:
            if col != "Year":
                df[col] = df[col].apply(lambda v: f"${v:,.0f}" if isinstance(v, (int, float)) else v)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("Pro forma data not available.")

    # Amortization schedule
    amort = pro_forma.get("amortization", [])
    if amort:
        st.markdown("---")
        st.subheader("Loan Amortization Schedule")
        amort_df = pd.DataFrame(amort)
        if "year" in amort_df.columns:
            dollar_cols = [c for c in amort_df.columns if c != "year"]
            for col in dollar_cols:
                amort_df[col] = amort_df[col].apply(
                    lambda v: f"${v:,.0f}" if isinstance(v, (int, float)) else v)
        st.dataframe(amort_df, use_container_width=True, hide_index=True)

    # Sensitivity tables
    if sensitivity and not isinstance(sensitivity.get("exit_cap"), str):
        st.markdown("---")
        st.subheader("Sensitivity: IRR vs Exit Cap Rate")
        exit_cap_tbl = sensitivity.get("exit_cap")
        if exit_cap_tbl and isinstance(exit_cap_tbl, dict):
            try:
                ec_df = pd.DataFrame(exit_cap_tbl)
                st.dataframe(ec_df, use_container_width=True)
            except Exception:
                st.json(exit_cap_tbl)


# ── Tab 3: Market Data ────────────────────────────────────────────────────────
with tabs[3]:
    st.subheader("Market Intelligence")
    demo = market.get("demographics", {}).get("structured", {}) if market else {}
    caps = market.get("cap_rates", {}) if market else {}
    crime = market.get("crime", {}) if market else {}
    ws = market.get("walkscore", {}) if market else {}
    hud = market.get("hud_fmr", {}) if market else {}
    rc = market.get("rentcast", {}) if market else {}
    fz = market.get("flood_zone", {}) if market else {}
    rc_stats = rc.get("market_stats", {}) if rc else {}
    rc_avm = rc.get("rent_estimate", {}) if rc else {}

    # Demographics
    if demo:
        st.markdown("**Demographics**")
        d1, d2, d3, d4, d5 = st.columns(5)
        d1.metric("Population", f"{demo.get('population', 0):,}" if demo.get("population") else "—")
        d2.metric("Median Income", _dollar(demo.get("median_income")))
        d3.metric("Median Rent", _dollar(demo.get("median_rent")))
        d4.metric("Unemployment", f"{demo.get('unemployment_rate', 0):.1f}%" if demo.get("unemployment_rate") else "—")
        d5.metric("Renter-Occupied", f"{demo.get('renter_pct', '—')}%")

    st.markdown("---")

    # Cap rates & rates
    if caps.get("average_cap_rate") or caps.get("treasury_10yr"):
        st.markdown("**Macro Rates (FRED)**")
        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("Market Cap Rate", f"{caps.get('average_cap_rate', 0):.2f}%"
                   if caps.get("average_cap_rate") else "—")
        rc2.metric("10-Yr Treasury", f"{caps.get('treasury_10yr', 0):.2f}%"
                   if caps.get("treasury_10yr") else "—")
        rc3.metric("30-Yr Mortgage", f"{caps.get('mortgage_30yr', 0):.2f}%"
                   if caps.get("mortgage_30yr") else "—")

    st.markdown("---")
    col_left, col_right = st.columns(2)

    # RentCast
    with col_left:
        st.markdown("**RentCast Market Comps**")
        if rc_stats.get("available"):
            st.write(f"Average Rent: **${rc_stats.get('average_rent', '—'):,}/mo**")
            st.write(f"Median Rent: **${rc_stats.get('median_rent', '—'):,}/mo**")
            st.write(f"Avg Days on Market: **{rc_stats.get('avg_days_on_market', '—')}**")
        else:
            st.caption("RentCast market stats unavailable (API key not configured).")

        if rc_avm.get("available") and rc_avm.get("estimated_rent"):
            st.success(
                f"RentCast AVM: **${rc_avm['estimated_rent']:,.0f}/mo** "
                f"(range: ${rc_avm.get('rent_range_low', '?'):,}–${rc_avm.get('rent_range_high', '?'):,})"
            )

        # HUD FMR
        st.markdown("**HUD Fair Market Rents**")
        if hud.get("available"):
            fmr = hud.get("fmr_by_bedroom", {})
            hud_cols = st.columns(4)
            for col, (br, key) in zip(hud_cols, [
                ("Studio", "efficiency"), ("1BR", "one_br"), ("2BR", "two_br"), ("3BR", "three_br")
            ]):
                col.metric(br, f"${fmr.get(key, '—'):,}" if fmr.get(key) else "—")
            st.caption(f"HUD Metro: {hud.get('metro_name', '—')}")
        else:
            st.caption("HUD FMR unavailable.")

    # Walk Score & Flood Zone & Crime
    with col_right:
        st.markdown("**Walk / Transit / Bike Scores**")
        if ws.get("available"):
            w1, w2, w3 = st.columns(3)
            w1.metric("Walk", f"{ws.get('walk_score', '—')} ({ws.get('walk_label', '')})")
            w2.metric("Transit", f"{ws.get('transit_score', '—')} ({ws.get('transit_label', '')})")
            w3.metric("Bike", f"{ws.get('bike_score', '—')} ({ws.get('bike_label', '')})")
        else:
            st.caption("Walk Score unavailable (API key not configured).")

        st.markdown("**FEMA Flood Zone**")
        if fz.get("available"):
            flood_risk = fz.get("risk_level", "Unknown")
            if "Low" in flood_risk:
                st.success(f"Flood Zone: {fz.get('flood_zone', '—')} — {flood_risk}")
            elif "High" in flood_risk or "Moderate" in flood_risk:
                st.warning(f"Flood Zone: {fz.get('flood_zone', '—')} — {flood_risk}")
            else:
                st.info(f"Flood Zone: {fz.get('flood_zone', '—')} — {flood_risk}")
            sfha = "Yes ⚠️" if fz.get("sfha") else "No"
            st.caption(f"SFHA (flood insurance required): {sfha}")
            st.caption(fz.get("description", ""))
        else:
            st.caption("FEMA data unavailable.")

        st.markdown("**Crime**")
        if crime.get("available"):
            risk = crime.get("risk_level", "Unknown")
            vs_nat = crime.get("vs_national_pct", 0)
            if "Low" in risk:
                st.success(f"Crime Risk: {risk} ({vs_nat:+.1f}% vs national)")
            elif "High" in risk or "Very High" in risk:
                st.error(f"Crime Risk: {risk} ({vs_nat:+.1f}% vs national)")
            else:
                st.warning(f"Crime Risk: {risk} ({vs_nat:+.1f}% vs national)")
            st.caption(f"Violent crimes/100k: {crime.get('violent_per_100k', '—')}")
        else:
            st.caption("Crime data unavailable.")


# ── Tab 4: ML & Monte Carlo ───────────────────────────────────────────────────
with tabs[4]:
    col_ml, col_mc = st.columns(2)

    with col_ml:
        st.subheader("ML Property Valuation")
        if ml and ml.get("available") and not ml.get("error"):
            assessment = ml.get("assessment", "UNKNOWN")
            disc = ml.get("premium_discount_pct", 0)
            if assessment == "UNDERVALUED":
                st.success(f"**{assessment}** — {abs(disc):.1f}% below predicted value")
            elif assessment == "OVERVALUED":
                st.error(f"**{assessment}** — {disc:.1f}% above predicted value")
            else:
                st.info(f"**{assessment}** — within fair value range")

            ml1, ml2 = st.columns(2)
            ml1.metric("Predicted Value",   _dollar(ml.get("predicted_value")))
            ml2.metric("Actual Price",      _dollar(ml.get("actual_price")))
            ml1.metric("Model R²",          _fmt(ml.get("test_r2")))
            ml2.metric("Premium/Discount",  f"{disc:+.1f}%")

            features = ml.get("top_features", [])
            if features:
                st.markdown("**Top Valuation Drivers**")
                for feat in features[:5]:
                    st.caption(f"• {feat}")
        elif ml.get("error"):
            st.warning(f"ML valuation unavailable: {ml['error']}")
        else:
            st.info("Enable ML valuation on the input form to see property value assessment.")

    with col_mc:
        st.subheader("Monte Carlo Simulation")
        if mc and mc.get("irr_mean") is not None and not mc.get("error"):
            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("P10 IRR", _pct(mc.get("irr_p10")))
            mc2.metric("P50 IRR", _pct(mc.get("irr_p50")))
            mc3.metric("P90 IRR", _pct(mc.get("irr_p90")))
            mc1.metric("Mean IRR", _pct(mc.get("irr_mean")))
            mc2.metric("Std Dev",  _pct(mc.get("irr_std")))

            probs = mc.get("probabilities", {})
            if probs:
                st.markdown("**Probability Table**")
                import pandas as pd
                prob_df = pd.DataFrame(
                    [{"Scenario": k, "Probability": f"{v:.1f}%"} for k, v in probs.items()])
                st.dataframe(prob_df, use_container_width=True, hide_index=True)

            # Histogram
            irr_samples = mc.get("irr_samples", [])
            if irr_samples:
                try:
                    import plotly.express as px
                    fig = px.histogram(
                        x=[x * 100 for x in irr_samples],
                        nbins=50,
                        title="IRR Distribution (1,000 simulations)",
                        labels={"x": "IRR (%)", "y": "Count"},
                        color_discrete_sequence=["#C9A84C"],
                    )
                    fig.update_layout(
                        paper_bgcolor="#0D1117",
                        plot_bgcolor="#161B22",
                        font_color="#E6EDF3",
                        title_font_color="#C9A84C",
                    )
                    st.plotly_chart(fig, use_container_width=True)
                except ImportError:
                    st.caption("Install plotly for distribution chart.")
        elif mc.get("error"):
            st.warning(f"Monte Carlo failed: {mc['error']}")
        else:
            st.info("Monte Carlo simulation did not run (check logs).")


# ── Tab 5: Rent Forecast ──────────────────────────────────────────────────────
with tabs[5]:
    st.subheader("Rent Growth Prediction")
    if rent_pred and not rent_pred.get("error"):
        import pandas as pd
        avg_growth = rent_pred.get("avg_predicted_growth", 0)
        st.metric("Average Predicted Growth", f"{avg_growth:.2f}%/yr")

        yr_table = rent_pred.get("yearly_predictions", [])
        if yr_table:
            rp_df = pd.DataFrame(yr_table)
            st.dataframe(rp_df, use_container_width=True, hide_index=True)

            # Plotly line chart
            try:
                import plotly.graph_objects as go
                fig = go.Figure()
                if "year" in rp_df.columns and "predicted_rent" in rp_df.columns:
                    fig.add_trace(go.Scatter(
                        x=rp_df["year"],
                        y=rp_df["predicted_rent"],
                        mode="lines+markers",
                        name="Predicted Rent",
                        line=dict(color="#C9A84C", width=2),
                    ))
                fig.update_layout(
                    title="Rent Prediction (Year-by-Year)",
                    xaxis_title="Year", yaxis_title="Rent ($/unit/mo)",
                    paper_bgcolor="#0D1117", plot_bgcolor="#161B22",
                    font_color="#E6EDF3",
                )
                st.plotly_chart(fig, use_container_width=True)
            except ImportError:
                pass

        # Backtest
        if backtest and not backtest.get("error"):
            st.markdown("---")
            st.subheader("Model Backtest Results")
            bt_cols = st.columns(3)
            bt_cols[0].metric("MAE",  f"{backtest.get('mae', 0):.2f}%")
            bt_cols[1].metric("RMSE", f"{backtest.get('rmse', 0):.2f}%")
            bt_cols[2].metric("R²",   _fmt(backtest.get("r2")))
    elif rent_pred.get("error"):
        st.warning(f"Rent prediction unavailable: {rent_pred['error']}")
    else:
        st.info("Enable Rent Growth Prediction on the input form to see ML-based rent forecasts.")


# ── Tab 6: Scenarios ──────────────────────────────────────────────────────────
with tabs[6]:
    st.subheader("Base / Bull / Bear Scenario Comparison")
    if scenarios and not scenarios.get("error"):
        import pandas as pd
        sc_rows = []
        for scenario_name, sc_data in scenarios.items():
            if isinstance(sc_data, dict):
                m = sc_data.get("metrics", {})
                sc_rows.append({
                    "Scenario":        scenario_name,
                    "Levered IRR":     _pct(m.get("levered_irr")),
                    "Equity Multiple": _x(m.get("equity_multiple")),
                    "Cash-on-Cash Y1": _pct(m.get("cash_on_cash_yr1")),
                    "DSCR Y1":         _fmt(m.get("dscr_yr1")),
                    "Exit Cap Rate":   _pct(m.get("exit_cap_rate")),
                })
        if sc_rows:
            sc_df = pd.DataFrame(sc_rows)
            st.dataframe(sc_df, use_container_width=True, hide_index=True)

        # NOI growth chart
        try:
            import plotly.graph_objects as go
            fig = go.Figure()
            colors = {"Base": "#C9A84C", "Bull": "#3fb950", "Bear": "#f85149"}
            for scenario_name, sc_data in scenarios.items():
                if isinstance(sc_data, dict):
                    nois = sc_data.get("annual_nois", [])
                    if nois:
                        fig.add_trace(go.Scatter(
                            x=list(range(1, len(nois) + 1)),
                            y=nois,
                            mode="lines+markers",
                            name=scenario_name,
                            line=dict(color=colors.get(scenario_name, "#8B949E"), width=2),
                        ))
            if fig.data:
                fig.update_layout(
                    title="NOI Growth by Scenario",
                    xaxis_title="Year", yaxis_title="NOI ($)",
                    paper_bgcolor="#0D1117", plot_bgcolor="#161B22",
                    font_color="#E6EDF3",
                )
                st.plotly_chart(fig, use_container_width=True)
        except ImportError:
            pass
    elif scenarios.get("error"):
        st.warning(f"Scenarios unavailable: {scenarios['error']}")
    else:
        st.info("Scenario comparison did not run.")


# ── Tab 7: Lease Analysis ─────────────────────────────────────────────────────
with tabs[7]:
    st.subheader("Lease Document Analysis")
    if lease and not lease.get("error"):
        if "portfolio_summary" in lease:
            ps = lease["portfolio_summary"]
            st.metric("Leases Analyzed", str(lease.get("lease_count", "—")))
            lc1, lc2 = st.columns(2)
            lc1.metric("Total Monthly Rent", _dollar(ps.get("total_monthly_rent")))
            lc2.metric("Weighted Avg Escalation", f"{ps.get('weighted_avg_escalation', 0):.1f}%/yr")

            risk_flags = ps.get("risk_flags", [])
            if risk_flags:
                st.warning("**Risk Flags**")
                for flag in risk_flags:
                    st.markdown(f"• {flag}")

            indiv = lease.get("individual_leases", [])
            if indiv:
                st.markdown("**Individual Leases**")
                import pandas as pd
                lease_rows = []
                for l in indiv:
                    lease_rows.append({
                        "File":         l.get("filename", "—"),
                        "Monthly Rent": _dollar(l.get("monthly_rent")),
                        "Term":         l.get("lease_term", "—"),
                        "Escalation":   f"{l.get('annual_escalation_pct', 0):.1f}%/yr",
                        "Expiry":       l.get("lease_expiry_date", "—"),
                        "Risk Flags":   ", ".join(l.get("risk_flags", [])) or "None",
                    })
                st.dataframe(pd.DataFrame(lease_rows), use_container_width=True, hide_index=True)
        else:
            # Single lease
            lc1, lc2, lc3 = st.columns(3)
            lc1.metric("Monthly Rent",     _dollar(lease.get("monthly_rent")))
            lc2.metric("Lease Term",       str(lease.get("lease_term", "—")))
            lc3.metric("Annual Escalation", f"{lease.get('annual_escalation_pct', 0):.1f}%/yr")
            st.metric("Expiry Date", str(lease.get("lease_expiry_date", "—")))

            risk_flags = lease.get("risk_flags", [])
            if risk_flags:
                st.warning("**Risk Flags**")
                for flag in risk_flags:
                    st.markdown(f"• {flag}")

        # Input comparison flags
        comp = lease.get("input_comparison", {})
        comp_flags = comp.get("flags", [])
        if comp_flags:
            st.markdown("---")
            st.warning("**Input Comparison Flags**")
            for flag in comp_flags:
                st.markdown(f"• {flag}")
    elif lease.get("error"):
        st.warning(f"Lease analysis unavailable: {lease['error']}")
    else:
        st.info("Upload lease PDFs on the input form to enable AI lease analysis.")


# ── Tab 8: AI Memo ────────────────────────────────────────────────────────────
with tabs[8]:
    st.subheader("AI Investment Memo")
    if ai_memo and not ai_memo.get("error"):
        sections = [
            ("executive_summary",   "Executive Summary"),
            ("market_analysis",     "Market Analysis"),
            ("financial_analysis",  "Financial Analysis"),
            ("risk_factors",        "Risk Factors"),
            ("investment_thesis",   "Investment Thesis"),
            ("recommendation",      "Recommendation"),
        ]
        for key, title in sections:
            content = ai_memo.get(key, "")
            if content:
                with st.expander(f"**{title}**", expanded=(key == "executive_summary")):
                    st.markdown(content)
    elif ai_memo.get("error"):
        st.warning(f"AI memo unavailable: {ai_memo['error']}")
    else:
        st.info("AI memo requires an OpenAI API key. Add OPENAI_API_KEY to your environment or Streamlit secrets.")


# ── Tab 9: Waterfall ──────────────────────────────────────────────────────────
with tabs[9]:
    st.subheader("LP/GP Waterfall Distribution")
    if waterfall and not waterfall.get("error"):
        import pandas as pd
        wf1, wf2 = st.columns(2)
        wf1.metric("LP IRR",         _pct(waterfall.get("lp_irr")))
        wf2.metric("GP IRR",         _pct(waterfall.get("gp_irr")))
        wf1.metric("LP Total Return", _dollar(waterfall.get("lp_total_return")))
        wf2.metric("GP Total Return", _dollar(waterfall.get("gp_total_return")))
        wf1.metric("LP Multiple",    _x(waterfall.get("lp_multiple")))
        wf2.metric("GP Multiple",    _x(waterfall.get("gp_multiple")))

        tiers = waterfall.get("tiers", [])
        if tiers:
            st.markdown("---")
            st.markdown("**Distribution Tiers**")
            tier_rows = []
            for t in tiers:
                tier_rows.append({
                    "Tier":        t.get("name", "—"),
                    "Total Dist":  _dollar(t.get("total_distribution")),
                    "LP Share":    _dollar(t.get("lp_distribution")),
                    "GP Share":    _dollar(t.get("gp_distribution")),
                    "LP %":        f"{t.get('lp_pct', 0):.0f}%",
                    "GP %":        f"{t.get('gp_pct', 0):.0f}%",
                })
            st.dataframe(pd.DataFrame(tier_rows), use_container_width=True, hide_index=True)
    elif waterfall and waterfall.get("error"):
        st.warning(f"Waterfall unavailable: {waterfall['error']}")
    else:
        st.info("Enable LP/GP Waterfall on the input form to see distribution analysis.")


# ── Tab 10: Downloads ─────────────────────────────────────────────────────────
with tabs[10]:
    st.subheader("Download Reports")
    dl1, dl2, dl3 = st.columns(3)

    with dl1:
        st.markdown("**📊 Excel Workbook**")
        st.caption("Full pro forma, sensitivity tables, market data, and ML results.")
        if excel_path and os.path.exists(excel_path):
            with open(excel_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Excel",
                    data=f,
                    file_name=f"underwriting_{R.get('job_id', 'report')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        else:
            st.caption("Excel not available.")

    with dl2:
        st.markdown("**📝 Word Memo**")
        st.caption("Investment memo with narrative, comps, and AI analysis.")
        if word_path and os.path.exists(word_path):
            with open(word_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Word",
                    data=f,
                    file_name=f"investment_memo_{R.get('job_id', 'report')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
        else:
            st.caption("Word document not available.")

    with dl3:
        st.markdown("**📄 PDF Report**")
        st.caption("Printable PDF summary of all underwriting findings.")
        if pdf_path and os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                st.download_button(
                    "⬇️ Download PDF",
                    data=f,
                    file_name=f"investment_report_{R.get('job_id', 'report')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
        else:
            st.caption("PDF not available.")
