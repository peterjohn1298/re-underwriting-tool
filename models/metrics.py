import numpy_financial as npf


def calc_irr(cash_flows: list[float]) -> float | None:
    """Calculate IRR from a list of cash flows (Year 0 negative, then annual)."""
    try:
        result = npf.irr(cash_flows)
        if result is None or result != result:  # NaN check
            return None
        return float(result)
    except Exception:
        return None


def calc_equity_multiple(cash_flows: list[float]) -> float:
    """Total distributions / total equity invested."""
    invested = abs(cash_flows[0])
    if invested == 0:
        return 0.0
    total_returned = sum(cash_flows[1:])
    return total_returned / invested


def calc_cash_on_cash(annual_cf: float, equity_invested: float) -> float:
    """Annual before-tax cash flow / equity invested."""
    if equity_invested == 0:
        return 0.0
    return annual_cf / equity_invested


def calc_dscr(noi: float, annual_debt_service: float) -> float:
    """Net Operating Income / Annual Debt Service."""
    if annual_debt_service == 0:
        return 0.0
    return noi / annual_debt_service


def calc_yield_on_cost(noi: float, total_cost: float) -> float:
    """Stabilized NOI / Total Project Cost."""
    if total_cost == 0:
        return 0.0
    return noi / total_cost


def calc_monthly_payment(principal: float, annual_rate: float, amort_years: int) -> float:
    """Calculate monthly mortgage payment (P&I)."""
    if principal == 0 or annual_rate == 0:
        return 0.0
    return float(-npf.pmt(annual_rate / 12, amort_years * 12, principal))


def build_amortization_schedule(
    principal: float, annual_rate: float, amort_years: int, term_years: int, io_years: int = 0
) -> list[dict]:
    """Build monthly amortization schedule."""
    monthly_rate = annual_rate / 12
    total_months = term_years * 12
    amort_months = amort_years * 12
    io_months = io_years * 12

    if principal == 0:
        return []

    pi_payment = calc_monthly_payment(principal, annual_rate, amort_years)
    io_payment = principal * monthly_rate

    schedule = []
    balance = principal

    for month in range(1, total_months + 1):
        if month <= io_months:
            interest = balance * monthly_rate
            principal_paid = 0.0
            payment = io_payment
        else:
            interest = balance * monthly_rate
            principal_paid = pi_payment - interest
            payment = pi_payment

        balance -= principal_paid
        schedule.append({
            "month": month,
            "payment": round(payment, 2),
            "principal": round(principal_paid, 2),
            "interest": round(interest, 2),
            "balance": round(max(balance, 0), 2),
        })

    return schedule


def sensitivity_table_exit_cap(
    base_noi_at_exit: float,
    base_exit_cap: float,
    equity_invested: float,
    annual_cfs: list[float],
    sale_costs_pct: float,
    loan_balance_at_exit: float,
    increments: list[float] | None = None,
) -> list[dict]:
    """Sensitivity on exit cap rate — returns rows with IRR and equity multiple."""
    if increments is None:
        increments = [-0.0100, -0.0050, -0.0025, 0.0, 0.0025, 0.0050, 0.0100]

    results = []
    for delta in increments:
        cap = base_exit_cap + delta
        if cap <= 0:
            continue
        sale_price = base_noi_at_exit / cap
        net_proceeds = sale_price * (1 - sale_costs_pct) - loan_balance_at_exit
        cfs = [-equity_invested] + list(annual_cfs) + [annual_cfs[-1] + net_proceeds] if annual_cfs else [-equity_invested, net_proceeds]
        # Rebuild: year 0 = equity, years 1..N-1 = annual CFs, year N = last CF + sale
        cfs_full = [-equity_invested] + list(annual_cfs[:-1]) + [annual_cfs[-1] + net_proceeds] if len(annual_cfs) > 0 else [-equity_invested, net_proceeds]
        irr = calc_irr(cfs_full)
        em = calc_equity_multiple(cfs_full)
        results.append({
            "exit_cap_rate": round(cap * 100, 2),
            "sale_price": round(sale_price),
            "net_proceeds": round(net_proceeds),
            "irr": round(irr * 100, 2) if irr else None,
            "equity_multiple": round(em, 2),
        })
    return results


def sensitivity_table_noi_growth(
    deal,
    growth_deltas: list[float] | None = None,
) -> list[dict]:
    """Sensitivity on NOI growth — returns rows with IRR and equity multiple."""
    if growth_deltas is None:
        growth_deltas = [-0.02, -0.01, 0.0, 0.01, 0.02]

    import copy
    from models.financial_model import build_pro_forma
    results = []
    for delta in growth_deltas:
        modified = copy.deepcopy(deal)
        modified.revenue_growth_rate = deal.revenue_growth_rate + delta
        pf = build_pro_forma(modified)
        irr = pf["metrics"]["levered_irr"]
        em = pf["metrics"]["equity_multiple"]
        results.append({
            "noi_growth_rate": round((deal.revenue_growth_rate + delta) * 100, 2),
            "irr": round(irr * 100, 2) if irr else None,
            "equity_multiple": round(em, 2) if em else None,
        })
    return results


def sensitivity_table_interest_rate(
    deal,
    rate_deltas: list[float] | None = None,
) -> list[dict]:
    """Sensitivity on interest rate — returns rows with IRR and DSCR."""
    if rate_deltas is None:
        rate_deltas = [-0.0150, -0.0100, -0.0050, 0.0, 0.0050, 0.0100, 0.0150]

    import copy
    from models.financial_model import build_pro_forma
    results = []
    for delta in rate_deltas:
        modified = copy.deepcopy(deal)
        modified.interest_rate = max(0.01, deal.interest_rate + delta)
        pf = build_pro_forma(modified)
        irr = pf["metrics"]["levered_irr"]
        em = pf["metrics"]["equity_multiple"]
        dscr = pf["metrics"]["dscr_yr1"]
        results.append({
            "interest_rate": round((deal.interest_rate + delta) * 100, 2),
            "irr": round(irr * 100, 2) if irr else None,
            "equity_multiple": round(em, 2) if em else None,
            "dscr": round(dscr, 2) if dscr else None,
        })
    return results


def sensitivity_table_rent_growth(
    deal,
    growth_deltas: list[float] | None = None,
) -> list[dict]:
    """Sensitivity on rent growth — returns rows with IRR and stabilized YOC."""
    if growth_deltas is None:
        growth_deltas = [-0.02, -0.01, 0.0, 0.01, 0.02]

    import copy
    from models.financial_model import build_pro_forma
    results = []
    for delta in growth_deltas:
        modified = copy.deepcopy(deal)
        modified.revenue_growth_rate = max(0, deal.revenue_growth_rate + delta)
        # Clear variable growth so flat rate is used
        modified.yearly_revenue_growth = []
        pf = build_pro_forma(modified)
        irr = pf["metrics"]["levered_irr"]
        em = pf["metrics"]["equity_multiple"]
        syoc = pf["metrics"]["stabilized_yoc"]
        results.append({
            "rent_growth": round((deal.revenue_growth_rate + delta) * 100, 2),
            "irr": round(irr * 100, 2) if irr else None,
            "equity_multiple": round(em, 2) if em else None,
            "stabilized_yoc": round(syoc * 100, 2) if syoc else None,
        })
    return results


def sensitivity_table_purchase_price(
    deal,
    price_deltas: list[float] | None = None,
) -> list[dict]:
    """Sensitivity on purchase price — returns rows with IRR and equity multiple."""
    if price_deltas is None:
        price_deltas = [-0.10, -0.05, 0.0, 0.05, 0.10]

    import copy
    from models.financial_model import build_pro_forma
    results = []
    for delta in price_deltas:
        modified = copy.deepcopy(deal)
        modified.purchase_price = deal.purchase_price * (1 + delta)
        pf = build_pro_forma(modified)
        irr = pf["metrics"]["levered_irr"]
        em = pf["metrics"]["equity_multiple"]
        results.append({
            "price_change": f"{delta:+.0%}",
            "purchase_price": round(modified.purchase_price),
            "irr": round(irr * 100, 2) if irr else None,
            "equity_multiple": round(em, 2) if em else None,
        })
    return results
