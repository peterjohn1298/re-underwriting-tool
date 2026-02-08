"""10-year pro forma engine — derives everything from DealInputs."""

from models.assumptions import DealInputs, DerivedAssumptions, derive_assumptions, to_full_dict
from models.metrics import (
    calc_irr, calc_equity_multiple, calc_cash_on_cash,
    calc_dscr, calc_yield_on_cost, calc_monthly_payment,
    build_amortization_schedule,
)


def build_pro_forma(deal: DealInputs) -> dict:
    """Build a 10-year pro forma from deal inputs. Returns all results."""
    derived = derive_assumptions(deal)
    years = 10
    hold = deal.hold_period_years

    # --- Sources & Uses ---
    closing_costs = deal.purchase_price * deal.closing_costs_pct
    sources_uses = {
        "sources": {
            "senior_debt": round(derived.loan_amount, 2),
            "sponsor_equity": round(derived.equity_required, 2),
            "total": round(derived.loan_amount + derived.equity_required, 2),
        },
        "uses": {
            "purchase_price": round(deal.purchase_price, 2),
            "closing_costs": round(closing_costs, 2),
            "deferred_maintenance": round(deal.deferred_maintenance, 2),
            "planned_capex": round(deal.planned_capex, 2),
            "total_capex": round(derived.total_capex, 2),
            "total": round(derived.total_project_cost, 2),
        },
    }

    # --- Debt Service ---
    loan_amount = derived.loan_amount
    annual_ds = 0.0
    monthly_payment = 0.0
    amort_schedule = []
    if loan_amount > 0 and deal.interest_rate > 0:
        monthly_payment = calc_monthly_payment(
            loan_amount, deal.interest_rate, deal.amortization_years
        )
        annual_ds = monthly_payment * 12
        amort_schedule = build_amortization_schedule(
            loan_amount, deal.interest_rate,
            deal.amortization_years, deal.loan_term_years,
            deal.io_period_years,
        )

    # --- Value-Add Rent Schedule ---
    # If market_rent > in_place_rent, ramp rents toward market over first 3 years
    in_place = deal.in_place_rent
    market = deal.market_rent if deal.market_rent > 0 else in_place
    rent_gap = market - in_place
    ramp_years = min(3, hold)  # ramp to market over 3 years

    # --- Annual Pro Forma ---
    pro_forma_years = []
    annual_nois = []
    annual_btcfs = []

    for yr in range(1, years + 1):
        # Rent per unit: ramp from in-place to market, then grow
        if rent_gap > 0 and yr <= ramp_years:
            # Linear ramp toward market rent
            rent_per_unit = in_place + rent_gap * (yr / ramp_years)
        elif rent_gap > 0:
            # Past ramp: grow from market rent
            rent_per_unit = market * (1 + deal.revenue_growth_rate) ** (yr - ramp_years)
        else:
            # No value-add: just grow from in-place
            rent_per_unit = in_place * (1 + deal.revenue_growth_rate) ** (yr - 1)

        gpr = rent_per_unit * deal.total_units * 12

        # Occupancy improvement toward 95% stabilized if capex is being done
        if derived.total_capex > 0 and deal.occupancy < 0.95:
            stabilized_occ = min(0.95, deal.occupancy + 0.03)
            if yr <= 2:
                occ = deal.occupancy + (stabilized_occ - deal.occupancy) * (yr / 2)
            else:
                occ = stabilized_occ
        else:
            occ = deal.occupancy

        vacancy_loss = gpr * (1 - occ)
        exp_growth = (1 + deal.expense_growth_rate) ** (yr - 1)

        # Other income grows with revenue
        other_inc = derived.other_income * (1 + deal.revenue_growth_rate) ** (yr - 1)
        egi = gpr * occ + other_inc

        # Expenses from derived breakdown, grown annually
        mgmt_fee = egi * deal.management_fee_pct
        prop_tax = derived.property_tax * exp_growth
        insurance = derived.insurance * exp_growth
        utilities = derived.utilities * exp_growth
        repairs = derived.repairs_maintenance * exp_growth
        ga = derived.general_admin * exp_growth
        other_exp = derived.other_expenses * exp_growth
        reserves = deal.replacement_reserves_per_unit * deal.total_units * exp_growth

        total_expenses = mgmt_fee + prop_tax + insurance + utilities + repairs + ga + other_exp + reserves
        noi = egi - total_expenses

        # Debt service
        if yr <= deal.io_period_years:
            yr_ds = loan_amount * deal.interest_rate
        else:
            yr_ds = annual_ds

        btcf = noi - yr_ds

        row = {
            "year": yr,
            "rent_per_unit": round(rent_per_unit, 0),
            "occupancy": round(occ, 4),
            "gross_potential_rent": round(gpr, 2),
            "vacancy_loss": round(vacancy_loss, 2),
            "other_income": round(other_inc, 2),
            "effective_gross_income": round(egi, 2),
            "management_fee": round(mgmt_fee, 2),
            "property_tax": round(prop_tax, 2),
            "insurance": round(insurance, 2),
            "utilities": round(utilities, 2),
            "repairs_maintenance": round(repairs, 2),
            "general_admin": round(ga, 2),
            "other_expenses": round(other_exp, 2),
            "replacement_reserves": round(reserves, 2),
            "total_expenses": round(total_expenses, 2),
            "noi": round(noi, 2),
            "debt_service": round(yr_ds, 2),
            "btcf": round(btcf, 2),
        }
        pro_forma_years.append(row)
        annual_nois.append(noi)
        annual_btcfs.append(btcf)

    # --- Exit / Reversion ---
    exit_year = min(hold, years)
    exit_noi = annual_nois[exit_year - 1]
    if exit_year < len(annual_nois):
        forward_noi = annual_nois[exit_year]
    else:
        forward_noi = exit_noi * (1 + deal.revenue_growth_rate)

    exit_cap = derived.exit_cap_rate
    if exit_cap <= 0:
        exit_cap = 0.06

    sale_price = forward_noi / exit_cap
    sale_costs = sale_price * deal.sale_costs_pct

    if amort_schedule and exit_year * 12 <= len(amort_schedule):
        loan_balance_at_exit = amort_schedule[exit_year * 12 - 1]["balance"]
    else:
        loan_balance_at_exit = loan_amount

    net_sale_proceeds = sale_price - sale_costs - loan_balance_at_exit

    reversion = {
        "exit_year": exit_year,
        "forward_noi": round(forward_noi, 2),
        "exit_cap_rate": round(exit_cap, 4),
        "sale_price": round(sale_price, 2),
        "sale_costs": round(sale_costs, 2),
        "loan_balance": round(loan_balance_at_exit, 2),
        "net_sale_proceeds": round(net_sale_proceeds, 2),
    }

    # --- Return Metrics ---
    equity = derived.equity_required
    total_cost = derived.total_project_cost

    levered_cfs = [-equity] + annual_btcfs[:exit_year - 1] + [annual_btcfs[exit_year - 1] + net_sale_proceeds]
    unlevered_cfs = [-total_cost] + annual_nois[:exit_year - 1] + [annual_nois[exit_year - 1] + sale_price - sale_costs]

    levered_irr = calc_irr(levered_cfs)
    unlevered_irr = calc_irr(unlevered_cfs)
    equity_mult = calc_equity_multiple(levered_cfs)
    coc_yr1 = calc_cash_on_cash(annual_btcfs[0], equity) if equity > 0 else 0
    dscr_yr1 = calc_dscr(annual_nois[0], annual_ds) if annual_ds > 0 else 0
    yoc = calc_yield_on_cost(annual_nois[0], total_cost)

    # Stabilized metrics (year 3)
    stabilized_noi = annual_nois[min(2, len(annual_nois) - 1)]
    stabilized_yoc = calc_yield_on_cost(stabilized_noi, total_cost)
    stabilized_dscr = calc_dscr(stabilized_noi, annual_ds) if annual_ds > 0 else 0

    metrics = {
        "levered_irr": levered_irr,
        "unlevered_irr": unlevered_irr,
        "equity_multiple": equity_mult,
        "cash_on_cash_yr1": coc_yr1,
        "dscr_yr1": dscr_yr1,
        "yield_on_cost": yoc,
        "stabilized_yoc": stabilized_yoc,
        "stabilized_dscr": stabilized_dscr,
        "going_in_cap_rate": derived.going_in_cap_rate,
        "exit_cap_rate": exit_cap,
        "price_per_unit": derived.price_per_unit,
        "price_per_sf": derived.price_per_sf,
    }

    # Annual amortization summary
    annual_amort = []
    if amort_schedule:
        for yr in range(1, min(years, len(amort_schedule) // 12) + 1):
            start = (yr - 1) * 12
            end = yr * 12
            yr_slice = amort_schedule[start:end]
            annual_amort.append({
                "year": yr,
                "total_payment": round(sum(m["payment"] for m in yr_slice), 2),
                "total_principal": round(sum(m["principal"] for m in yr_slice), 2),
                "total_interest": round(sum(m["interest"] for m in yr_slice), 2),
                "ending_balance": yr_slice[-1]["balance"],
            })

    inputs_dict = to_full_dict(deal, derived)

    return {
        "inputs": inputs_dict,
        "sources_uses": sources_uses,
        "pro_forma": pro_forma_years,
        "amortization_monthly": amort_schedule,
        "amortization_annual": annual_amort,
        "reversion": reversion,
        "metrics": metrics,
        "levered_cash_flows": [round(cf, 2) for cf in levered_cfs],
        "unlevered_cash_flows": [round(cf, 2) for cf in unlevered_cfs],
        "annual_nois": [round(n, 2) for n in annual_nois],
        "annual_btcfs": [round(b, 2) for b in annual_btcfs],
    }
