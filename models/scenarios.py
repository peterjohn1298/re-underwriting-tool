"""Base / Bull / Bear scenario comparison.

Runs three variants of the financial model side-by-side:
  Base: user's original assumptions
  Bull: higher rent growth, better occupancy, tighter exit cap
  Bear: lower rent growth, weaker occupancy, wider exit cap

All offsets are expressed as absolute deltas from the Base case.
"""

import copy
import logging

logger = logging.getLogger(__name__)

# Default offsets (can be overridden via form or API)
BULL_OFFSETS = {
    "revenue_growth_rate": +0.015,    # +1.5% higher rent growth
    "occupancy": +0.03,               # +3% better occupancy
    "exit_cap_rate_spread": -0.0025,  # exit cap 25bps tighter
}

BEAR_OFFSETS = {
    "revenue_growth_rate": -0.015,    # -1.5% lower rent growth
    "occupancy": -0.05,               # -5% weaker occupancy
    "exit_cap_rate_spread": +0.005,   # exit cap 50bps wider
}


def run_scenarios(deal, build_pro_forma_fn, bull_offsets: dict = None, bear_offsets: dict = None) -> dict:
    """Run Base, Bull, and Bear scenarios and return a comparison dict.

    Args:
        deal: DealInputs instance
        build_pro_forma_fn: the build_pro_forma function
        bull_offsets: optional override for Bull scenario offsets
        bear_offsets: optional override for Bear scenario offsets

    Returns:
        dict with keys: base, bull, bear, comparison_table
    """
    bull = bull_offsets or BULL_OFFSETS
    bear = bear_offsets or BEAR_OFFSETS

    results = {}

    for label, offsets in [("base", {}), ("bull", bull), ("bear", bear)]:
        try:
            sim_deal = copy.deepcopy(deal)
            sim_deal.yearly_revenue_growth = []  # use flat rate for scenarios

            for attr, delta in offsets.items():
                current = getattr(sim_deal, attr, None)
                if current is not None:
                    new_val = current + delta
                    # Clamp occupancy to [0.5, 1.0]
                    if attr == "occupancy":
                        new_val = max(0.5, min(1.0, new_val))
                    setattr(sim_deal, attr, new_val)

            pf = build_pro_forma_fn(sim_deal)
            m = pf["metrics"]
            rev = pf["reversion"]
            results[label] = {
                "label": label.title(),
                "revenue_growth_rate": sim_deal.revenue_growth_rate,
                "occupancy": sim_deal.occupancy,
                "exit_cap_rate_spread": sim_deal.exit_cap_rate_spread,
                "levered_irr": m.get("levered_irr"),
                "equity_multiple": m.get("equity_multiple"),
                "dscr_yr1": m.get("dscr_yr1"),
                "cash_on_cash_yr1": m.get("cash_on_cash_yr1"),
                "going_in_cap_rate": m.get("going_in_cap_rate"),
                "exit_cap_rate": m.get("exit_cap_rate"),
                "sale_price": rev.get("sale_price"),
                "net_sale_proceeds": rev.get("net_sale_proceeds"),
                "year1_noi": pf["pro_forma"][0]["noi"] if pf.get("pro_forma") else None,
                "year1_btcf": pf["pro_forma"][0]["btcf"] if pf.get("pro_forma") else None,
            }
        except Exception as e:
            logger.error(f"Scenario {label} failed: {e}")
            results[label] = {"label": label.title(), "error": str(e)}

    # Build comparison table rows
    def _pct(v):
        return f"{v * 100:.1f}%" if v is not None else "N/A"

    def _cur(v):
        if v is None:
            return "N/A"
        if abs(v) >= 1_000_000:
            return f"${v / 1_000_000:.1f}M"
        return f"${v:,.0f}"

    def _x(v):
        return f"{v:.2f}x" if v is not None else "N/A"

    comparison_rows = [
        ("Rent Growth (p.a.)", lambda s: _pct(s.get("revenue_growth_rate"))),
        ("Occupancy", lambda s: _pct(s.get("occupancy"))),
        ("Exit Cap Rate", lambda s: _pct(s.get("exit_cap_rate"))),
        ("Yr 1 NOI", lambda s: _cur(s.get("year1_noi"))),
        ("Yr 1 BTCF", lambda s: _cur(s.get("year1_btcf"))),
        ("DSCR (Yr 1)", lambda s: _x(s.get("dscr_yr1"))),
        ("Cash-on-Cash (Yr 1)", lambda s: _pct(s.get("cash_on_cash_yr1"))),
        ("Levered IRR", lambda s: _pct(s.get("levered_irr"))),
        ("Equity Multiple", lambda s: _x(s.get("equity_multiple"))),
        ("Exit Sale Price", lambda s: _cur(s.get("sale_price"))),
        ("Net Sale Proceeds", lambda s: _cur(s.get("net_sale_proceeds"))),
    ]

    table = []
    for metric, fmt_fn in comparison_rows:
        row = {"metric": metric}
        for label in ("base", "bull", "bear"):
            s = results.get(label, {})
            row[label] = fmt_fn(s) if not s.get("error") else "Error"
        table.append(row)

    return {
        "base": results.get("base", {}),
        "bull": results.get("bull", {}),
        "bear": results.get("bear", {}),
        "comparison_table": table,
        "bull_offsets": bull,
        "bear_offsets": bear,
    }
