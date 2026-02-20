"""LP/GP Waterfall / Promote Model.

Standard two-tier promote waterfall:
  Tier 0: Return of capital (LP and GP each get their equity back)
  Tier 1: LP preferred return (compounded annually on LP equity)
  Tier 2: Residual split — (1 - promote_pct) to LP, promote_pct to GP

Inputs
------
lp_equity_pct      : LP share of total equity (e.g., 0.90 → 90%)
preferred_return   : LP annualized preferred return (e.g., 0.08 → 8%)
promote_pct        : GP's promote share of residual above preferred (e.g., 0.20 → 80/20)
equity_invested    : total equity at deal close (from DerivedAssumptions.equity_required)
annual_btcfs       : list of annual before-tax cash flows (after debt service), years 1-N
exit_proceeds      : net sale proceeds after loan payoff (and tax if after-tax)
hold_years         : hold period in integer years

Outputs
-------
All dollar amounts in the result dict are distributions (positive = cash received).
LP and GP IRR are computed from their respective net cash flow streams.
"""

import logging

logger = logging.getLogger(__name__)


def run_waterfall(
    lp_equity_pct: float,
    preferred_return: float,
    promote_pct: float,
    equity_invested: float,
    annual_btcfs: list,
    exit_proceeds: float,
    hold_years: int,
) -> dict:
    """Run the LP/GP waterfall and return a complete distribution summary."""
    if equity_invested <= 0:
        return {"error": "equity_invested must be positive"}
    if not (0 < lp_equity_pct < 1):
        return {"error": "lp_equity_pct must be between 0 and 1 (exclusive)"}

    lp_equity = equity_invested * lp_equity_pct
    gp_equity = equity_invested * (1 - lp_equity_pct)

    # --- Total cash available to distribute ---
    operating_total = sum(cf for cf in annual_btcfs if cf > 0)
    exit_total = max(0.0, exit_proceeds)
    total_available = operating_total + exit_total

    # --- Waterfall tier allocation (aggregate basis) ---
    remaining = total_available

    # Tier 0: Return of LP capital
    lp_return_of_capital = min(lp_equity, remaining)
    remaining -= lp_return_of_capital

    # Tier 0: Return of GP capital
    gp_return_of_capital = min(gp_equity, remaining)
    remaining -= gp_return_of_capital

    # Tier 1: LP preferred return (compounded) on original LP equity
    lp_pref_due = lp_equity * ((1 + preferred_return) ** hold_years - 1)
    lp_pref_paid = min(lp_pref_due, remaining)
    lp_pref_unpaid = lp_pref_due - lp_pref_paid
    remaining -= lp_pref_paid

    # Tier 2: Residual promote split
    lp_residual = remaining * (1 - promote_pct)
    gp_promote = remaining * promote_pct
    remaining = 0.0

    # --- Totals ---
    lp_total = lp_return_of_capital + lp_pref_paid + lp_residual
    gp_total = gp_return_of_capital + gp_promote

    lp_em = lp_total / lp_equity if lp_equity > 0 else None
    gp_em = gp_total / gp_equity if gp_equity > 0 else None

    # --- Per-year LP / GP cash flows for IRR calculation ---
    # Operating CFs: distribute pro-rata (pre-waterfall, no promote on current income)
    lp_annual_cfs = [-lp_equity]
    gp_annual_cfs = [-gp_equity]

    for cf in annual_btcfs:
        lp_annual_cfs.append(cf * lp_equity_pct)
        gp_annual_cfs.append(cf * (1 - lp_equity_pct))

    # Exit year: allocate waterfall residual to LP/GP, adjusted for operating already paid
    operating_paid_to_lp = sum(cf * lp_equity_pct for cf in annual_btcfs if cf > 0)
    operating_paid_to_gp = sum(cf * (1 - lp_equity_pct) for cf in annual_btcfs if cf > 0)
    lp_exit_cf = lp_total - operating_paid_to_lp
    gp_exit_cf = gp_total - operating_paid_to_gp

    lp_annual_cfs.append(lp_exit_cf)
    gp_annual_cfs.append(gp_exit_cf)

    lp_irr = _safe_irr(lp_annual_cfs)
    gp_irr = _safe_irr(gp_annual_cfs)

    # --- Promote economics summary ---
    gp_capital_return_pct = (gp_total / gp_equity - 1) * 100 if gp_equity > 0 else 0
    promote_dollars = gp_promote  # how much GP earned from promote vs their equity contribution
    promote_on_gp_equity = promote_dollars / gp_equity if gp_equity > 0 else 0

    # --- Build per-tier waterfall table ---
    tiers = [
        {
            "tier": "Return of LP Capital",
            "lp_amount": lp_return_of_capital,
            "gp_amount": 0.0,
            "note": f"LP equity returned ({lp_equity_pct*100:.0f}% of total)",
        },
        {
            "tier": "Return of GP Capital",
            "lp_amount": 0.0,
            "gp_amount": gp_return_of_capital,
            "note": f"GP equity returned ({(1-lp_equity_pct)*100:.0f}% of total)",
        },
        {
            "tier": f"LP Preferred Return ({preferred_return*100:.1f}% p.a.)",
            "lp_amount": lp_pref_paid,
            "gp_amount": 0.0,
            "note": f"Due: ${lp_pref_due:,.0f}" + (" (fully paid)" if lp_pref_unpaid < 1 else f" (${lp_pref_unpaid:,.0f} unpaid)"),
        },
        {
            "tier": f"Residual — LP {(1-promote_pct)*100:.0f}% / GP {promote_pct*100:.0f}%",
            "lp_amount": lp_residual,
            "gp_amount": gp_promote,
            "note": "Above-preferred cash flow (promote)",
        },
        {
            "tier": "TOTAL",
            "lp_amount": lp_total,
            "gp_amount": gp_total,
            "note": f"Total distributions: ${total_available:,.0f}",
        },
    ]

    return {
        # Inputs
        "lp_equity": lp_equity,
        "gp_equity": gp_equity,
        "lp_equity_pct": lp_equity_pct,
        "gp_equity_pct": 1 - lp_equity_pct,
        "preferred_return": preferred_return,
        "promote_pct": promote_pct,
        "hold_years": hold_years,
        # Distribution summary
        "lp_return_of_capital": lp_return_of_capital,
        "lp_pref_due": lp_pref_due,
        "lp_pref_paid": lp_pref_paid,
        "lp_pref_unpaid": lp_pref_unpaid,
        "lp_residual": lp_residual,
        "lp_total": lp_total,
        "gp_return_of_capital": gp_return_of_capital,
        "gp_promote": gp_promote,
        "gp_total": gp_total,
        # Return metrics
        "lp_irr": lp_irr,
        "gp_irr": gp_irr,
        "lp_equity_multiple": lp_em,
        "gp_equity_multiple": gp_em,
        "promote_dollars": promote_dollars,
        "promote_on_gp_equity": promote_on_gp_equity,
        "gp_capital_return_pct": gp_capital_return_pct,
        # Meta
        "total_distributions": total_available,
        "pref_fully_satisfied": lp_pref_unpaid < 1.0,
        "tiers": tiers,
    }


def _safe_irr(cash_flows: list) -> float | None:
    """Compute IRR, returning None on failure."""
    try:
        import numpy_financial as npf
        result = npf.irr(cash_flows)
        if result is None or result != result:  # NaN check
            return None
        return float(result)
    except Exception:
        return None
