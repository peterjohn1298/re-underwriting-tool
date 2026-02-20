"""Streamlit entry point — RE Underwriting Intelligence Platform."""

import os
import sys
import uuid
import copy
import logging
from pathlib import Path

import streamlit as st

# ── Project root on sys.path ──────────────────────────────────────────────────
sys.path.insert(0, str(Path(__file__).parent))

from config import OUTPUT_DIR, UPLOAD_DIR

# ── Page config (must be first Streamlit call) ────────────────────────────────
st.set_page_config(
    page_title="RE Underwriting Platform",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── CSS — navy/gold dark theme ────────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background-color: #0D1117; }
[data-testid="stHeader"] { background-color: #0D1117; }
.stTabs [data-baseweb="tab-list"] { gap: 6px; background: #0D1117; }
.stTabs [data-baseweb="tab"] {
    background-color: #161B22;
    border-radius: 6px;
    padding: 8px 18px;
    color: #8B949E;
    font-weight: 500;
}
.stTabs [aria-selected="true"] {
    background-color: #C9A84C !important;
    color: #0D1117 !important;
    font-weight: bold;
}
div[data-testid="stForm"] {
    background: #161B22;
    border-radius: 10px;
    padding: 20px;
    border: 1px solid #30363D;
}
.section-header {
    color: #C9A84C;
    font-size: 1.1em;
    font-weight: 700;
    margin-top: 12px;
    margin-bottom: 4px;
    border-bottom: 1px solid #30363D;
    padding-bottom: 4px;
}
</style>
""", unsafe_allow_html=True)

# ── Helper: resolve API key from secrets → env ────────────────────────────────

def _secret(key: str, fallback: str = "") -> str:
    try:
        return st.secrets.get(key, fallback) or fallback
    except Exception:
        return fallback


# ── Demo values (Austin, TX — 100-unit Class B) ───────────────────────────────
DEMO_VALUES = {
    "property_type":        "Multifamily - Class B",
    "address":              "2100 S Lamar Blvd, Austin, TX 78704",
    "year_built":           2005,
    "purchase_price":       12_500_000.0,
    "current_noi":          750_000.0,
    "total_units":          100,
    "total_sf":             85_000.0,
    "in_place_rent":        1_450.0,
    "market_rent":          1_600.0,
    "occupancy":            94,
    "hold_period_years":    7,
    "deferred_maintenance": 250_000.0,
    "planned_capex":        500_000.0,
    "capex_description":    "Unit renovations and lobby refresh",
    # Unit mix
    "um_count_Studio": 0,   "um_rent_Studio": 0.0,   "um_market_rent_Studio": 0.0,
    "um_count_1BR":    40,  "um_rent_1BR":    1_250.0, "um_market_rent_1BR":  1_400.0,
    "um_count_2BR":    50,  "um_rent_2BR":    1_600.0, "um_market_rent_2BR":  1_750.0,
    "um_count_3BR":    10,  "um_rent_3BR":    1_900.0, "um_market_rent_3BR":  2_050.0,
    # Financing / overrides (all at default)
    "ltv":                        65,
    "interest_rate":              6.75,
    "amortization_years":         30,
    "loan_term_years":            10,
    "io_period_years":            0,
    "closing_costs_pct":          3.0,
    "sale_costs_pct":             2.5,
    "replacement_reserves_per_unit": 250.0,
    "revenue_growth_rate":        3.0,
    "expense_growth_rate":        3.0,
    "management_fee_pct":         3.5,
    "exit_cap_rate_spread":       25,
    "tax_rate":                   25.0,
    "land_value_pct":             20.0,
    "lp_equity_pct":              90.0,
    "preferred_return":           8.0,
    "promote_pct":                20.0,
    "override_property_tax":      0.0,
    "override_insurance":         0.0,
    "override_utilities":         0.0,
    "override_repairs":           0.0,
    "override_general_admin":     0.0,
    "override_other_expenses":    0.0,
}

# ── Demo values #2 (Phoenix, AZ — 80-unit Value-Add Class C) ─────────────────
DEMO_VALUES_2 = {
    "property_type":        "Multifamily - Class C",
    "address":              "4500 N 12th St, Phoenix, AZ 85014",
    "year_built":           1987,
    "purchase_price":       7_200_000.0,
    "current_noi":          0.0,
    "total_units":          80,
    "total_sf":             56_000.0,
    "in_place_rent":        850.0,
    "market_rent":          1_200.0,
    "occupancy":            78,
    "hold_period_years":    5,
    "deferred_maintenance": 600_000.0,
    "planned_capex":        2_000_000.0,
    "capex_description":    "Full unit gut-renovations, new roofs, HVAC replacement, parking reseal",
    # Unit mix
    "um_count_Studio": 10,  "um_rent_Studio": 750.0,  "um_market_rent_Studio": 1_050.0,
    "um_count_1BR":    40,  "um_rent_1BR":    850.0,  "um_market_rent_1BR":   1_200.0,
    "um_count_2BR":    25,  "um_rent_2BR":    950.0,  "um_market_rent_2BR":   1_350.0,
    "um_count_3BR":     5,  "um_rent_3BR":  1_100.0,  "um_market_rent_3BR":   1_550.0,
    # Financing / overrides
    "ltv":                        60,
    "interest_rate":              7.25,
    "amortization_years":         30,
    "loan_term_years":            10,
    "io_period_years":            2,
    "closing_costs_pct":          3.5,
    "sale_costs_pct":             3.0,
    "replacement_reserves_per_unit": 300.0,
    "revenue_growth_rate":        4.5,
    "expense_growth_rate":        3.5,
    "management_fee_pct":         4.0,
    "exit_cap_rate_spread":       0,
    "tax_rate":                   25.0,
    "land_value_pct":             15.0,
    "lp_equity_pct":              90.0,
    "preferred_return":           8.0,
    "promote_pct":                25.0,
    "override_property_tax":      0.0,
    "override_insurance":         0.0,
    "override_utilities":         0.0,
    "override_repairs":           0.0,
    "override_general_admin":     0.0,
    "override_other_expenses":    0.0,
}

# ── Demo values #3 (Miami, FL — 150-unit Class A Core) ───────────────────────
DEMO_VALUES_3 = {
    "property_type":        "Multifamily - Class A",
    "address":              "1800 Brickell Ave, Miami, FL 33129",
    "year_built":           2018,
    "purchase_price":       45_000_000.0,
    "current_noi":          0.0,
    "total_units":          150,
    "total_sf":             165_000.0,
    "in_place_rent":        2_800.0,
    "market_rent":          2_950.0,
    "occupancy":            96,
    "hold_period_years":    10,
    "deferred_maintenance": 0.0,
    "planned_capex":        150_000.0,
    "capex_description":    "Common area refresh and smart-home upgrades",
    # Unit mix
    "um_count_Studio": 20,  "um_rent_Studio": 2_200.0, "um_market_rent_Studio": 2_400.0,
    "um_count_1BR":    70,  "um_rent_1BR":   2_800.0, "um_market_rent_1BR":   2_950.0,
    "um_count_2BR":    50,  "um_rent_2BR":   3_400.0, "um_market_rent_2BR":   3_600.0,
    "um_count_3BR":    10,  "um_rent_3BR":   4_200.0, "um_market_rent_3BR":   4_400.0,
    # Financing / overrides
    "ltv":                        55,
    "interest_rate":              6.25,
    "amortization_years":         30,
    "loan_term_years":            10,
    "io_period_years":            0,
    "closing_costs_pct":          2.5,
    "sale_costs_pct":             2.5,
    "replacement_reserves_per_unit": 200.0,
    "revenue_growth_rate":        3.5,
    "expense_growth_rate":        3.0,
    "management_fee_pct":         3.0,
    "exit_cap_rate_spread":       25,
    "tax_rate":                   25.0,
    "land_value_pct":             30.0,
    "lp_equity_pct":              90.0,
    "preferred_return":           7.0,
    "promote_pct":                20.0,
    "override_property_tax":      0.0,
    "override_insurance":         0.0,
    "override_utilities":         0.0,
    "override_repairs":           0.0,
    "override_general_admin":     0.0,
    "override_other_expenses":    0.0,
}

PROPERTY_TYPES = [
    "Multifamily - Class A",
    "Multifamily - Class B",
    "Multifamily - Class C",
    "Mixed-Use - Class B",
    "Student Housing - Class B",
    "Senior Housing - Class B",
]

# ── Helper functions (defined before form so they are available on submit) ────

def _compare_lease_to_inputs(lease_data: dict, deal) -> dict:
    comparison = {"flags": []}
    if "individual_leases" in lease_data:
        total_monthly = lease_data.get("portfolio_summary", {}).get("total_monthly_rent")
        if total_monthly and deal.in_place_rent > 0 and deal.total_units > 0:
            expected_total = deal.in_place_rent * deal.total_units
            diff_pct = abs(total_monthly - expected_total) / expected_total * 100
            if diff_pct > 10:
                comparison["flags"].append(
                    f"Lease portfolio total rent (${total_monthly:,.0f}/mo) differs from "
                    f"underwriting assumption (${expected_total:,.0f}/mo) by {diff_pct:.0f}%"
                )
            comparison["lease_total_monthly"] = total_monthly
            comparison["input_total_monthly"] = expected_total
    else:
        lease_rent = lease_data.get("monthly_rent")
        if lease_rent and deal.in_place_rent > 0:
            diff_pct = abs(lease_rent - deal.in_place_rent) / deal.in_place_rent * 100
            if diff_pct > 10:
                comparison["flags"].append(
                    f"Lease rent (${lease_rent:,.0f}/mo) differs from in-place rent "
                    f"(${deal.in_place_rent:,.0f}/mo) by {diff_pct:.0f}%"
                )
        escalation = lease_data.get("annual_escalation_pct")
        if escalation and deal.revenue_growth_rate > 0:
            growth_pct = deal.revenue_growth_rate * 100
            if abs(escalation - growth_pct) > 1.0:
                comparison["flags"].append(
                    f"Lease escalation ({escalation}%/yr) vs assumed growth ({growth_pct:.1f}%/yr)"
                )
    return comparison


def _build_sensitivity(deal, pro_forma: dict) -> dict:
    from models.metrics import (
        sensitivity_table_exit_cap, sensitivity_table_interest_rate,
        sensitivity_table_rent_growth, sensitivity_table_purchase_price,
    )
    rev = pro_forma["reversion"]
    derived = pro_forma["inputs"]["derived"]
    return {
        "exit_cap": sensitivity_table_exit_cap(
            base_noi_at_exit=rev["forward_noi"],
            base_exit_cap=rev["exit_cap_rate"],
            equity_invested=derived["equity_required"],
            annual_cfs=pro_forma["annual_btcfs"][:rev["exit_year"]],
            sale_costs_pct=deal.sale_costs_pct,
            loan_balance_at_exit=rev["loan_balance"],
        ),
        "interest_rate": sensitivity_table_interest_rate(deal),
        "rent_growth": sensitivity_table_rent_growth(deal),
        "purchase_price": sensitivity_table_purchase_price(deal),
    }


def _build_recommendation(pro_forma: dict, ml_valuation=None,
                           lease_analysis=None, rent_prediction=None,
                           monte_carlo=None) -> dict:
    m = pro_forma["metrics"]
    irr = m.get("levered_irr")
    signals, score = [], 0

    if irr is not None:
        if irr >= 0.15:
            signals.append(("IRR", "STRONG BUY", f"{irr*100:.1f}% exceeds 15% threshold"))
            score += 2
        elif irr >= 0.12:
            signals.append(("IRR", "BUY", f"{irr*100:.1f}% exceeds 12% threshold"))
            score += 1
        elif irr >= 0.08:
            signals.append(("IRR", "HOLD", f"{irr*100:.1f}% — moderate returns"))
        else:
            signals.append(("IRR", "PASS", f"{irr*100:.1f}% below 8% minimum"))
            score -= 2

    dscr = m.get("dscr_yr1", 0)
    if dscr < 1.0:
        signals.append(("DSCR", "WARNING", f"{dscr:.2f}x — negative cash flow"))
        score -= 2
    elif dscr < 1.25:
        signals.append(("DSCR", "CAUTION", f"{dscr:.2f}x — thin coverage"))
        score -= 1
    elif dscr >= 1.5:
        signals.append(("DSCR", "POSITIVE", f"{dscr:.2f}x — strong debt coverage"))
        score += 1

    coc = m.get("cash_on_cash_yr1", 0) or 0
    if coc >= 0.07:
        signals.append(("Cash-on-Cash", "POSITIVE", f"{coc*100:.1f}% Yr 1 — strong current yield"))
        score += 1
    elif coc >= 0.05:
        signals.append(("Cash-on-Cash", "NEUTRAL", f"{coc*100:.1f}% Yr 1 — acceptable yield"))
    elif coc < 0:
        signals.append(("Cash-on-Cash", "WARNING", f"{coc*100:.1f}% Yr 1 — negative cash flow"))
        score -= 1

    if ml_valuation and not ml_valuation.get("error"):
        assessment = ml_valuation.get("assessment", "")
        discount = ml_valuation.get("premium_discount_pct", 0)
        if assessment == "UNDERVALUED":
            signals.append(("ML Valuation", "BUY", f"ML shows {abs(discount):.1f}% discount to predicted value"))
            score += 1
        elif assessment == "OVERVALUED":
            signals.append(("ML Valuation", "CAUTION", f"ML shows {discount:.1f}% premium to predicted value"))
            score -= 1
        else:
            signals.append(("ML Valuation", "NEUTRAL", "Price within fair value range"))

    if lease_analysis and not lease_analysis.get("error"):
        comp_flags = lease_analysis.get("input_comparison", {}).get("flags", [])
        if comp_flags:
            signals.append(("Lease Analysis", "WARNING", "; ".join(comp_flags)))
            score -= 1
        risk_flags = lease_analysis.get("risk_flags", [])
        if len(risk_flags) > 3:
            signals.append(("Lease Risk", "CAUTION", f"{len(risk_flags)} risk flags identified"))
            score -= 1

    if rent_prediction and not rent_prediction.get("error"):
        avg_growth = rent_prediction.get("avg_predicted_growth", 0)
        if avg_growth > 4.0:
            signals.append(("Rent Forecast", "POSITIVE", f"{avg_growth:.1f}% avg predicted growth"))
            score += 1
        elif avg_growth < 1.0:
            signals.append(("Rent Forecast", "CAUTION", f"{avg_growth:.1f}% predicted growth — weak outlook"))
            score -= 1

    if monte_carlo and not monte_carlo.get("error"):
        mc_signal = monte_carlo.get("mc_signal", "NEUTRAL")
        mc_detail = monte_carlo.get("mc_detail", "")
        if mc_signal == "POSITIVE":
            signals.append(("Monte Carlo", "POSITIVE", mc_detail))
            score += 1
        elif mc_signal == "WARNING":
            signals.append(("Monte Carlo", "WARNING", mc_detail))
            score -= 1
        else:
            signals.append(("Monte Carlo", "NEUTRAL", mc_detail))

    if score >= 2:
        rec = "STRONG BUY"
    elif score >= 1:
        rec = "BUY"
    elif score >= 0:
        rec = "HOLD / CONDITIONAL"
    else:
        rec = "PASS"

    return {
        "recommendation": rec,
        "score": score,
        "signals": signals,
        "used_variable_growth": m.get("used_variable_growth", False),
    }


# ── Header ────────────────────────────────────────────────────────────────────
st.title("🏢 RE Underwriting Intelligence Platform")
st.caption("Institutional-Grade Multifamily Analysis • Powered by AI & ML")

# ── Demo buttons (outside form so they can trigger rerun) ─────────────────────
col_hdr, col_d1, col_d2, col_d3 = st.columns([3, 1, 1, 1])

def _load_demo(vals: dict) -> None:
    """Helper: push a demo dict into session state and rerun."""
    st.session_state["demo"] = vals.copy()
    for _ut in ["Studio", "1BR", "2BR", "3BR"]:
        st.session_state[f"um_count_{_ut}"]       = vals[f"um_count_{_ut}"]
        st.session_state[f"um_rent_{_ut}"]         = vals[f"um_rent_{_ut}"]
        st.session_state[f"um_market_rent_{_ut}"]  = vals[f"um_market_rent_{_ut}"]
    st.session_state["ck_ml_valuation"]    = True
    st.session_state["ck_rent_prediction"] = True

with col_d1:
    if st.button("📋 Demo 1 — Austin TX", use_container_width=True,
                 help="100-unit Class B • $12.5M • Stabilised value-add"):
        _load_demo(DEMO_VALUES)
        st.rerun()

with col_d2:
    if st.button("🔨 Demo 2 — Phoenix AZ", use_container_width=True,
                 help="80-unit Class C • $7.2M • Distressed value-add with heavy capex"):
        _load_demo(DEMO_VALUES_2)
        st.rerun()

with col_d3:
    if st.button("🏙️ Demo 3 — Miami FL", use_container_width=True,
                 help="150-unit Class A • $45M • Core stabilised luxury"):
        _load_demo(DEMO_VALUES_3)
        st.rerun()

_demo_loaded = st.session_state.get("demo", {})
if _demo_loaded:
    _addr = _demo_loaded.get("address", "demo deal")
    _units = _demo_loaded.get("total_units", "?")
    _price = _demo_loaded.get("purchase_price", 0)
    st.info(f"✅ Demo loaded — **{_addr}** ({_units} units, ${_price:,.0f}). "
            "Click **Run Full Analysis** to proceed.")

st.markdown("---")

# ── Read current demo values (empty dict if not loaded) ──────────────────────
d = st.session_state.get("demo", {})

# ── Input Form ────────────────────────────────────────────────────────────────
with st.form("deal_form"):
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 Deal Sheet",
        "🏠 Unit Mix",
        "🤖 AI / ML Features",
        "⚙️ Override Defaults",
    ])

    # ── Tab 1: Deal Sheet ─────────────────────────────────────────────────────
    with tab1:
        st.markdown('<div class="section-header">Property Information</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            _ptype_default = d.get("property_type", "Multifamily - Class B")
            _ptype_idx = PROPERTY_TYPES.index(_ptype_default) if _ptype_default in PROPERTY_TYPES else 1
            property_type = st.selectbox("Property Type", PROPERTY_TYPES, index=_ptype_idx)
            address = st.text_input("Address",
                                    value=d.get("address", ""),
                                    placeholder="123 Main St, Austin, TX 78701")
            year_built = st.number_input("Year Built", min_value=1900, max_value=2030,
                                          value=d.get("year_built", 2000), step=1)
        with c2:
            purchase_price = st.number_input("Purchase Price ($)", min_value=0.0,
                                              value=d.get("purchase_price", 5_000_000.0),
                                              step=50_000.0, format="%.0f")
            current_noi = st.number_input("Current NOI ($/yr)", min_value=0.0,
                                           value=d.get("current_noi", 0.0),
                                           step=10_000.0, format="%.0f",
                                           help="Leave 0 to back-calculate from rent & occupancy")
            total_units = st.number_input("Total Units", min_value=0,
                                           value=d.get("total_units", 100), step=1)
        with c3:
            total_sf = st.number_input("Total Sq Ft", min_value=0.0,
                                        value=d.get("total_sf", 0.0),
                                        step=1_000.0, format="%.0f",
                                        help="Optional — used for $/SF metrics")
            in_place_rent = st.number_input("In-Place Rent ($/unit/mo)", min_value=0.0,
                                             value=d.get("in_place_rent", 1_200.0),
                                             step=25.0, format="%.0f")
            market_rent = st.number_input("Market Rent ($/unit/mo)", min_value=0.0,
                                           value=d.get("market_rent", 0.0),
                                           step=25.0, format="%.0f",
                                           help="Leave 0 if same as in-place rent")

        st.markdown('<div class="section-header">Operating Details</div>', unsafe_allow_html=True)
        c4, c5, c6 = st.columns(3)
        with c4:
            occupancy = st.slider("Occupancy (%)", min_value=50, max_value=100,
                                   value=d.get("occupancy", 92), step=1)
            hold_period_years = st.number_input("Hold Period (years)", min_value=1, max_value=30,
                                                 value=d.get("hold_period_years", 7), step=1)
        with c5:
            deferred_maintenance = st.number_input("Deferred Maintenance ($)", min_value=0.0,
                                                    value=d.get("deferred_maintenance", 0.0),
                                                    step=10_000.0, format="%.0f")
            planned_capex = st.number_input("Planned CapEx ($)", min_value=0.0,
                                             value=d.get("planned_capex", 0.0),
                                             step=10_000.0, format="%.0f")
        with c6:
            capex_description = st.text_area("CapEx Description",
                                              value=d.get("capex_description", ""),
                                              placeholder="Roof, unit renovations, HVAC...",
                                              height=88)

    # ── Tab 2: Unit Mix ───────────────────────────────────────────────────────
    with tab2:
        st.markdown('<div class="section-header">Unit Mix Breakdown (optional)</div>',
                    unsafe_allow_html=True)
        st.caption("If filled in, blended rents here override the in-place/market rent above. "
                   "Leave counts at 0 to skip.")
        um_cols = st.columns(4)
        unit_types = ["Studio", "1BR", "2BR", "3BR"]
        unit_mix_data: dict = {}
        for i, ut in enumerate(unit_types):
            with um_cols[i]:
                st.markdown(f"**{ut}**")
                unit_mix_data[f"um_count_{ut}"] = st.number_input(
                    "# Units", min_value=0,
                    value=d.get(f"um_count_{ut}", 0),
                    step=1, key=f"um_count_{ut}")
                unit_mix_data[f"um_rent_{ut}"] = st.number_input(
                    "In-Place Rent ($)", min_value=0.0,
                    value=d.get(f"um_rent_{ut}", 0.0),
                    step=25.0, format="%.0f", key=f"um_rent_{ut}")
                unit_mix_data[f"um_market_rent_{ut}"] = st.number_input(
                    "Market Rent ($)", min_value=0.0,
                    value=d.get(f"um_market_rent_{ut}", 0.0),
                    step=25.0, format="%.0f", key=f"um_market_rent_{ut}")

        # Live blended rent preview
        total_um_units = sum(unit_mix_data.get(f"um_count_{ut}", 0) for ut in unit_types)
        if total_um_units > 0:
            weighted_sum = sum(
                unit_mix_data.get(f"um_count_{ut}", 0) * unit_mix_data.get(f"um_rent_{ut}", 0.0)
                for ut in unit_types
            )
            blended_rent = weighted_sum / total_um_units if total_um_units > 0 else 0
            st.info(f"Blended in-place rent: **${blended_rent:,.0f}/unit/mo** across {total_um_units} units")

    # ── Tab 3: AI/ML Features ─────────────────────────────────────────────────
    with tab3:
        st.markdown('<div class="section-header">AI & ML Features</div>', unsafe_allow_html=True)
        c7, c8 = st.columns(2)
        with c7:
            enable_ml_valuation = st.checkbox(
                "Enable ML Property Valuation", key="ck_ml_valuation",
                help="GradientBoosting model trained on market comps to assess fair value")
            enable_rent_prediction = st.checkbox(
                "Enable Rent Growth Prediction", key="ck_rent_prediction",
                help="ML-based rent growth forecast using FRED CPI & ZORI data; feeds into pro forma")
            enable_waterfall = st.checkbox(
                "Enable LP/GP Waterfall", value=False,
                help="Calculate LP preferred return and GP promote from deal cash flows")
        with c8:
            lease_files = st.file_uploader(
                "Upload Lease PDFs (optional)",
                type=["pdf"],
                accept_multiple_files=True,
                help="AI extracts terms, flags risks, and compares rents to your inputs")

    # ── Tab 4: Override Defaults ──────────────────────────────────────────────
    with tab4:
        st.markdown('<div class="section-header">Financing</div>', unsafe_allow_html=True)
        c9, c10, c11 = st.columns(3)
        with c9:
            ltv = st.slider("LTV (%)", min_value=40, max_value=85,
                             value=d.get("ltv", 65), step=1)
            interest_rate = st.number_input("Interest Rate (%)", min_value=0.0, max_value=20.0,
                                             value=d.get("interest_rate", 6.75),
                                             step=0.25, format="%.2f")
            amortization_years = st.number_input("Amortization (years)", min_value=1, max_value=40,
                                                  value=d.get("amortization_years", 30), step=1)
        with c10:
            loan_term_years = st.number_input("Loan Term (years)", min_value=1, max_value=30,
                                               value=d.get("loan_term_years", 10), step=1)
            io_period_years = st.number_input("I/O Period (years)", min_value=0, max_value=10,
                                               value=d.get("io_period_years", 0), step=1)
            closing_costs_pct = st.number_input("Closing Costs (%)", min_value=0.0, max_value=10.0,
                                                 value=d.get("closing_costs_pct", 3.0),
                                                 step=0.25, format="%.2f")
        with c11:
            sale_costs_pct = st.number_input("Sale Costs (%)", min_value=0.0, max_value=10.0,
                                              value=d.get("sale_costs_pct", 2.5),
                                              step=0.25, format="%.2f")
            replacement_reserves_per_unit = st.number_input(
                "Reserves ($/unit/yr)", min_value=0.0,
                value=d.get("replacement_reserves_per_unit", 250.0),
                step=50.0, format="%.0f")

        st.markdown('<div class="section-header">Growth Assumptions</div>', unsafe_allow_html=True)
        c12, c13 = st.columns(2)
        with c12:
            revenue_growth_rate = st.number_input("Revenue Growth (%/yr)", min_value=0.0,
                                                   max_value=20.0,
                                                   value=d.get("revenue_growth_rate", 3.0),
                                                   step=0.25, format="%.2f")
            expense_growth_rate = st.number_input("Expense Growth (%/yr)", min_value=0.0,
                                                   max_value=20.0,
                                                   value=d.get("expense_growth_rate", 3.0),
                                                   step=0.25, format="%.2f")
            management_fee_pct = st.number_input("Management Fee (%)", min_value=0.0, max_value=15.0,
                                                   value=d.get("management_fee_pct", 3.5),
                                                   step=0.25, format="%.2f")
        with c13:
            exit_cap_rate_spread = st.number_input(
                "Exit Cap Spread (bps)", min_value=-100, max_value=200,
                value=d.get("exit_cap_rate_spread", 25), step=5,
                help="Basis points added to going-in cap rate at exit")
            tax_rate = st.number_input("Marginal Tax Rate (%)", min_value=0.0, max_value=50.0,
                                        value=d.get("tax_rate", 25.0),
                                        step=1.0, format="%.0f")
            land_value_pct = st.number_input("Land Value (%)", min_value=0.0, max_value=80.0,
                                              value=d.get("land_value_pct", 20.0),
                                              step=5.0, format="%.0f",
                                              help="Land % of purchase price (not depreciable)")

        st.markdown('<div class="section-header">Waterfall (LP/GP)</div>', unsafe_allow_html=True)
        c14, c15 = st.columns(2)
        with c14:
            lp_equity_pct = st.number_input("LP Equity (%)", min_value=0.0, max_value=100.0,
                                             value=d.get("lp_equity_pct", 90.0),
                                             step=5.0, format="%.0f")
            preferred_return = st.number_input("LP Preferred Return (%)", min_value=0.0,
                                                max_value=20.0,
                                                value=d.get("preferred_return", 8.0),
                                                step=0.5, format="%.1f")
        with c15:
            promote_pct = st.number_input("GP Promote (%)", min_value=0.0, max_value=50.0,
                                           value=d.get("promote_pct", 20.0),
                                           step=5.0, format="%.0f")

        st.markdown('<div class="section-header">Expense Overrides (leave at 0 to auto-derive)</div>',
                    unsafe_allow_html=True)
        c16, c17, c18 = st.columns(3)
        with c16:
            override_property_tax = st.number_input("Property Tax ($/yr)", min_value=0.0,
                                                     value=d.get("override_property_tax", 0.0),
                                                     step=5_000.0, format="%.0f")
            override_insurance = st.number_input("Insurance ($/yr)", min_value=0.0,
                                                  value=d.get("override_insurance", 0.0),
                                                  step=1_000.0, format="%.0f")
        with c17:
            override_utilities = st.number_input("Utilities ($/yr)", min_value=0.0,
                                                  value=d.get("override_utilities", 0.0),
                                                  step=1_000.0, format="%.0f")
            override_repairs = st.number_input("Repairs & Maint ($/yr)", min_value=0.0,
                                                value=d.get("override_repairs", 0.0),
                                                step=1_000.0, format="%.0f")
        with c18:
            override_general_admin = st.number_input("G&A ($/yr)", min_value=0.0,
                                                      value=d.get("override_general_admin", 0.0),
                                                      step=1_000.0, format="%.0f")
            override_other_expenses = st.number_input("Other Expenses ($/yr)", min_value=0.0,
                                                       value=d.get("override_other_expenses", 0.0),
                                                       step=1_000.0, format="%.0f")

    st.markdown("---")
    submitted = st.form_submit_button(
        "🚀 Run Full Analysis", use_container_width=True, type="primary")


# ── Pipeline execution ────────────────────────────────────────────────────────
if submitted:
    # --- Validation ---
    errors = []
    if purchase_price <= 0:
        errors.append("Purchase price is required")
    if total_units <= 0:
        errors.append("Number of units is required")
    if in_place_rent <= 0 and all(unit_mix_data.get(f"um_rent_{ut}", 0) == 0 for ut in unit_types):
        errors.append("In-place rent is required (or fill in the Unit Mix tab)")
    if not address.strip():
        errors.append("Address is required for market research (e.g. 123 Main St, Austin, TX 78701)")
    if errors:
        for e in errors:
            st.error(f"❌ {e}")
        st.stop()

    # --- Build DealInputs ---
    from models.assumptions import DealInputs
    from models.unit_mix import parse_unit_mix_from_dict
    from models.financial_model import build_pro_forma

    deal = DealInputs(
        property_type=property_type,
        address=address.strip(),
        year_built=int(year_built),
        purchase_price=float(purchase_price),
        current_noi=float(current_noi),
        total_units=int(total_units),
        total_sf=float(total_sf),
        in_place_rent=float(in_place_rent),
        market_rent=float(market_rent),
        occupancy=float(occupancy) / 100.0,
        deferred_maintenance=float(deferred_maintenance),
        planned_capex=float(planned_capex),
        capex_description=capex_description,
        hold_period_years=int(hold_period_years),
        ltv=float(ltv) / 100.0,
        interest_rate=float(interest_rate) / 100.0,
        amortization_years=int(amortization_years),
        loan_term_years=int(loan_term_years),
        io_period_years=int(io_period_years),
        closing_costs_pct=float(closing_costs_pct) / 100.0,
        revenue_growth_rate=float(revenue_growth_rate) / 100.0,
        expense_growth_rate=float(expense_growth_rate) / 100.0,
        management_fee_pct=float(management_fee_pct) / 100.0,
        exit_cap_rate_spread=float(exit_cap_rate_spread) / 10000.0,
        sale_costs_pct=float(sale_costs_pct) / 100.0,
        replacement_reserves_per_unit=float(replacement_reserves_per_unit),
        tax_rate=float(tax_rate) / 100.0,
        land_value_pct=float(land_value_pct) / 100.0,
        override_property_tax=float(override_property_tax),
        override_insurance=float(override_insurance),
        override_utilities=float(override_utilities),
        override_repairs=float(override_repairs),
        override_general_admin=float(override_general_admin),
        override_other_expenses=float(override_other_expenses),
        enable_ml_valuation=bool(enable_ml_valuation),
        enable_rent_prediction=bool(enable_rent_prediction),
        enable_waterfall=bool(enable_waterfall),
        lp_equity_pct=float(lp_equity_pct) / 100.0,
        preferred_return=float(preferred_return) / 100.0,
        promote_pct=float(promote_pct) / 100.0,
    )

    # Parse and apply unit mix
    unit_mix = parse_unit_mix_from_dict(unit_mix_data)
    if not unit_mix.is_empty():
        deal.unit_mix = unit_mix
        if deal.total_units == 0:
            deal.total_units = unit_mix.total_units
        if unit_mix.blended_in_place_rent > 0:
            deal.in_place_rent = round(unit_mix.blended_in_place_rent, 2)
        if unit_mix.blended_market_rent > 0:
            deal.market_rent = round(unit_mix.blended_market_rent, 2)

    # Save uploaded lease files
    saved_paths = []
    if lease_files:
        for f in lease_files:
            if f is not None:
                tmp_path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex[:8]}_{f.name}")
                with open(tmp_path, "wb") as fp:
                    fp.write(f.read())
                saved_paths.append(tmp_path)
    if saved_paths:
        deal.lease_pdf_paths = saved_paths

    # --- Run pipeline ---
    logger = logging.getLogger(__name__)
    progress = st.progress(0, text="Starting analysis...")
    results_dict: dict = {}

    try:
        from models.assumptions import derive_assumptions
        from services.market_research import run_full_research

        progress.progress(5, text="Deriving assumptions...")
        derived = derive_assumptions(deal)

        progress.progress(10, text="Fetching market data (FRED, Census, BLS, RentCast)...")
        avg_sq_ft = (int(deal.total_sf / deal.total_units)
                     if deal.total_units > 0 and deal.total_sf > 0 else None)
        market_data = run_full_research(
            derived.asset_type, derived.city, derived.state,
            address=deal.address,
            zip_code=derived.zip_code,
            avg_bedrooms=2,
            avg_sq_ft=avg_sq_ft,
        )
        results_dict["market_data"] = market_data

        rent_prediction = None
        backtest_result = None
        cpi_shelter = None
        if deal.enable_rent_prediction:
            progress.progress(20, text="Running rent growth prediction (ML)...")
            try:
                from models.rent_predictor import RentPredictor
                from services.api_clients.fred_client import FREDClient
                fred = FREDClient()
                cpi_shelter = fred.get_cpi_shelter(limit=60)
                zori_rates = market_data.get("rent_trends", {}).get("zori_growth_rates", [])
                predictor = RentPredictor()
                predictor.train(cpi_shelter.get("annual_growth_rates", []),
                                zori_growth_rates=zori_rates)
                rent_prediction = predictor.predict(
                    deal.hold_period_years, deal.in_place_rent, deal.total_units)
                if rent_prediction and not rent_prediction.get("error"):
                    predicted_rates = rent_prediction.get("predicted_rates", [])
                    if predicted_rates:
                        deal.yearly_revenue_growth = predicted_rates
                from models.backtest import RentBacktester
                backtester = RentBacktester()
                backtest_result = backtester.run_backtest(
                    cpi_shelter.get("annual_growth_rates", []),
                    zori_growth_rates=zori_rates,
                )
            except Exception as e:
                logger.error(f"Rent prediction failed: {e}")
                rent_prediction = {"error": str(e)}
        results_dict["rent_prediction"] = rent_prediction
        results_dict["backtest"] = backtest_result

        progress.progress(35, text="Building financial model (pro forma)...")
        pro_forma = build_pro_forma(deal)
        results_dict["results"] = pro_forma

        ml_valuation = None
        if deal.enable_ml_valuation:
            progress.progress(45, text="Running ML property valuation...")
            try:
                from models.ml_valuation import PropertyValuationModel
                ml_model = PropertyValuationModel()
                ml_model.train(market_data)
                ml_valuation = ml_model.predict(deal, derived, market_data)
            except Exception as e:
                logger.error(f"ML valuation failed: {e}")
                ml_valuation = {"error": str(e)}
        results_dict["ml_valuation"] = ml_valuation

        lease_analysis = None
        if deal.lease_pdf_paths:
            progress.progress(55, text="Analyzing lease documents (AI)...")
            try:
                from services.lease_analyzer import LeaseAnalyzer
                analyzer = LeaseAnalyzer()
                if len(deal.lease_pdf_paths) == 1:
                    lease_analysis = analyzer.analyze_lease(deal.lease_pdf_paths[0])
                else:
                    lease_analysis = analyzer.analyze_multiple_leases(deal.lease_pdf_paths)
                if lease_analysis and not lease_analysis.get("error"):
                    lease_analysis["input_comparison"] = _compare_lease_to_inputs(
                        lease_analysis, deal)
            except Exception as e:
                logger.error(f"Lease analysis failed: {e}")
                lease_analysis = {"error": str(e)}
        results_dict["lease_analysis"] = lease_analysis

        progress.progress(58, text="Detecting housing market regime (FRED)...")
        try:
            from models.regime_detector import HousingRegimeDetector
            regime = HousingRegimeDetector().detect(market_data)
        except Exception as e:
            logger.error(f"Regime detection failed: {e}")
            regime = {"error": str(e)}
        results_dict["regime"] = regime

        progress.progress(60, text="Building sensitivity tables...")
        try:
            sensitivity = _build_sensitivity(deal, pro_forma)
        except Exception as e:
            logger.error(f"Sensitivity failed: {e}")
            sensitivity = {"error": str(e)}
        results_dict["sensitivity"] = sensitivity

        progress.progress(65, text="Running Monte Carlo simulation (1,000 paths)...")
        monte_carlo = None
        try:
            from models.monte_carlo import MonteCarloSimulator
            mc_sim = MonteCarloSimulator(n_iterations=1000)
            monte_carlo = mc_sim.run(deal, build_pro_forma)
        except Exception as e:
            logger.error(f"Monte Carlo failed: {e}")
            monte_carlo = {"error": str(e)}
        results_dict["monte_carlo"] = monte_carlo

        progress.progress(70, text="Building integrated recommendation...")
        recommendation = _build_recommendation(
            pro_forma, ml_valuation, lease_analysis, rent_prediction, monte_carlo)
        results_dict["recommendation"] = recommendation

        waterfall = None
        if deal.enable_waterfall:
            progress.progress(73, text="Running LP/GP waterfall...")
            try:
                from models.waterfall import run_waterfall
                derived_wf = pro_forma["inputs"]["derived"]
                rev_wf = pro_forma["reversion"]
                waterfall = run_waterfall(
                    lp_equity_pct=deal.lp_equity_pct,
                    preferred_return=deal.preferred_return,
                    promote_pct=deal.promote_pct,
                    equity_invested=derived_wf["equity_required"],
                    annual_btcfs=pro_forma.get("annual_btcfs", [])[:deal.hold_period_years],
                    exit_proceeds=rev_wf.get("net_sale_proceeds", 0),
                    hold_years=deal.hold_period_years,
                )
            except Exception as e:
                logger.error(f"Waterfall failed: {e}")
                waterfall = {"error": str(e)}
        results_dict["waterfall"] = waterfall

        ai_memo = None
        openai_key = _secret("OPENAI_API_KEY") or os.environ.get("OPENAI_API_KEY", "")
        if openai_key:
            progress.progress(77, text="Generating AI investment memo (GPT-4o)...")
            try:
                from services.ai_memo import generate_ai_memo
                memo_job = {
                    "deal": deal,
                    "results": pro_forma,
                    "market_data": market_data,
                    "ml_valuation": ml_valuation,
                    "monte_carlo": monte_carlo,
                    "recommendation": recommendation,
                    "rent_prediction": rent_prediction,
                }
                ai_memo = generate_ai_memo(memo_job, openai_key)
            except Exception as e:
                logger.error(f"AI memo failed: {e}")
                ai_memo = {"error": str(e)}
        results_dict["ai_memo"] = ai_memo

        progress.progress(82, text="Running Base/Bull/Bear scenarios...")
        try:
            from models.scenarios import run_scenarios
            scenarios = run_scenarios(deal, build_pro_forma)
        except Exception as e:
            logger.error(f"Scenarios failed: {e}")
            scenarios = {"error": str(e)}
        results_dict["scenarios"] = scenarios

        unit_mix_out = None
        if deal.unit_mix and not deal.unit_mix.is_empty():
            hud_fmr = market_data.get("hud_fmr", {}) if market_data else {}
            unit_mix_out = deal.unit_mix.with_hud_fmr(hud_fmr)
        results_dict["unit_mix"] = unit_mix_out

        job_id = uuid.uuid4().hex[:12]

        progress.progress(87, text="Generating Excel report...")
        try:
            from services.excel_generator import generate_excel
            excel_path = generate_excel(
                pro_forma, market_data, job_id,
                ml_valuation=ml_valuation, lease_analysis=lease_analysis,
                rent_prediction=rent_prediction, sensitivity=sensitivity,
                backtest=backtest_result, monte_carlo=monte_carlo)
            results_dict["excel_path"] = excel_path
        except Exception as e:
            logger.error(f"Excel generation failed: {e}")
            results_dict["excel_path"] = None

        progress.progress(92, text="Generating Word memo...")
        try:
            from services.word_generator import generate_word
            word_path = generate_word(
                pro_forma, market_data, job_id,
                ml_valuation=ml_valuation, lease_analysis=lease_analysis,
                rent_prediction=rent_prediction, sensitivity=sensitivity,
                backtest=backtest_result, monte_carlo=monte_carlo,
                ai_memo=ai_memo, unit_mix=unit_mix_out)
            results_dict["word_path"] = word_path
        except Exception as e:
            logger.error(f"Word generation failed: {e}")
            results_dict["word_path"] = None

        progress.progress(96, text="Generating PDF report...")
        try:
            from services.pdf_generator import generate_pdf
            pdf_path = generate_pdf(
                pro_forma, market_data, job_id,
                ml_valuation=ml_valuation, lease_analysis=lease_analysis,
                rent_prediction=rent_prediction, sensitivity=sensitivity,
                backtest=backtest_result, monte_carlo=monte_carlo)
            results_dict["pdf_path"] = pdf_path
        except Exception as e:
            logger.error(f"PDF generation failed: {e}")
            results_dict["pdf_path"] = None

        results_dict["deal"] = deal
        results_dict["job_id"] = job_id

        progress.progress(100, text="Analysis complete!")

        st.session_state["results"] = results_dict
        st.session_state["chat_history"] = []
        if "saved_deals" not in st.session_state:
            st.session_state["saved_deals"] = {}
        # Clear demo state after successful run
        st.session_state.pop("demo", None)

        st.success("✅ Analysis complete!")
        st.switch_page("pages/1_Results.py")

    except Exception as e:
        progress.empty()
        st.error(f"❌ Analysis failed: {str(e)}")
        st.exception(e)
