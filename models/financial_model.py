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
    in_place = deal.in_place_rent
    market = deal.market_rent if deal.market_rent > 0 else in_place
    rent_gap = market - in_place
    ramp_years = min(3, hold)

    # --- Variable Growth Rates (Fix #1 & #6) ---
    # Use year-by-year rates from rent predictor if available, else flat rate
    yearly_growth = deal.yearly_revenue_growth
    has_variable_growth = len(yearly_growth) > 0

    def _get_revenue_growth(yr_idx):
        """Get revenue growth rate for a given year (0-indexed)."""
        if has_variable_growth and yr_idx < len(yearly_growth):
            return yearly_growth[yr_idx] / 100  # predictor returns percentages
        return deal.revenue_growth_rate

    # --- Tax / Depreciation ---
    annual_depreciation = derived.annual_depreciation
    tax_rate = deal.tax_rate

    # --- Annual Pro Forma ---
    pro_forma_years = []
    annual_nois = []
    annual_btcfs = []
    annual_atcfs = []  # after-tax cash flows

    for yr in range(1, years + 1):
        growth = _get_revenue_growth(yr - 1)

        # Rent per unit: ramp from in-place to market, then grow
        if rent_gap > 0 and yr <= ramp_years:
            rent_per_unit = in_place + rent_gap * (yr / ramp_years)
        elif rent_gap > 0:
            # Past ramp: grow from market rent using variable or flat rate
            if has_variable_growth:
                # Compound from market rent using each year's rate
                rent_per_unit = market
                for y in range(ramp_years, yr):
                    g = _get_revenue_growth(y)
                    rent_per_unit *= (1 + g)
            else:
                rent_per_unit = market * (1 + deal.revenue_growth_rate) ** (yr - ramp_years)
        else:
            if has_variable_growth:
                rent_per_unit = in_place
                for y in range(0, yr - 1):
                    g = _get_revenue_growth(y)
                    rent_per_unit *= (1 + g)
            else:
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
        if has_variable_growth:
            other_inc = derived.other_income
            for y in range(0, yr - 1):
                other_inc *= (1 + _get_revenue_growth(y))
        else:
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

        # --- Tax Analysis (Fix #9) ---
        # Interest portion of debt service
        if amort_schedule and yr * 12 <= len(amort_schedule):
            start_month = (yr - 1) * 12
            end_month = yr * 12
            yr_interest = sum(m["interest"] for m in amort_schedule[start_month:end_month])
        elif yr <= deal.io_period_years:
            yr_interest = yr_ds  # all interest during IO
        else:
            yr_interest = loan_amount * deal.interest_rate  # approximation

        taxable_income = noi - yr_interest - annual_depreciation
        tax_liability = max(0, taxable_income * tax_rate)
        atcf = btcf - tax_liability

        row = {
            "year": yr,
            "rent_per_unit": round(rent_per_unit, 0),
            "revenue_growth_used": round(growth * 100, 2),
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
            # Tax fields
            "interest_expense": round(yr_interest, 2),
            "depreciation": round(annual_depreciation, 2),
            "taxable_income": round(taxable_income, 2),
            "tax_liability": round(tax_liability, 2),
            "atcf": round(atcf, 2),
        }
        pro_forma_years.append(row)
        annual_nois.append(noi)
        annual_btcfs.append(btcf)
        annual_atcfs.append(atcf)

    # --- Exit / Reversion ---
    exit_year = min(hold, years)
    exit_noi = annual_nois[exit_year - 1]
    if exit_year < len(annual_nois):
        forward_noi = annual_nois[exit_year]
    else:
        forward_noi = exit_noi * (1 + _get_revenue_growth(exit_year))

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

    # Tax on sale (depreciation recapture + capital gains)
    total_depreciation_taken = annual_depreciation * exit_year
    adjusted_basis = derived.total_project_cost - total_depreciation_taken
    capital_gain = max(0, sale_price - sale_costs - adjusted_basis)
    depreciation_recapture = min(capital_gain, total_depreciation_taken) * 0.25  # 25% recapture rate
    remaining_gain = max(0, capital_gain - total_depreciation_taken) * 0.20  # 20% LTCG
    tax_on_sale = depreciation_recapture + remaining_gain
    net_sale_proceeds_after_tax = net_sale_proceeds - tax_on_sale

    reversion = {
        "exit_year": exit_year,
        "forward_noi": round(forward_noi, 2),
        "exit_cap_rate": round(exit_cap, 4),
        "sale_price": round(sale_price, 2),
        "sale_costs": round(sale_costs, 2),
        "loan_balance": round(loan_balance_at_exit, 2),
        "net_sale_proceeds": round(net_sale_proceeds, 2),
        # Tax fields
        "total_depreciation": round(total_depreciation_taken, 2),
        "adjusted_basis": round(adjusted_basis, 2),
        "capital_gain": round(capital_gain, 2),
        "depreciation_recapture_tax": round(depreciation_recapture, 2),
        "capital_gains_tax": round(remaining_gain, 2),
        "total_tax_on_sale": round(tax_on_sale, 2),
        "net_sale_proceeds_after_tax": round(net_sale_proceeds_after_tax, 2),
    }

    # --- Return Metrics ---
    equity = derived.equity_required
    total_cost = derived.total_project_cost

    # Before-tax
    levered_cfs = [-equity] + annual_btcfs[:exit_year - 1] + [annual_btcfs[exit_year - 1] + net_sale_proceeds]
    unlevered_cfs = [-total_cost] + annual_nois[:exit_year - 1] + [annual_nois[exit_year - 1] + sale_price - sale_costs]

    # After-tax
    levered_atcfs = [-equity] + annual_atcfs[:exit_year - 1] + [annual_atcfs[exit_year - 1] + net_sale_proceeds_after_tax]

    levered_irr = calc_irr(levered_cfs)
    unlevered_irr = calc_irr(unlevered_cfs)
    after_tax_irr = calc_irr(levered_atcfs)
    equity_mult = calc_equity_multiple(levered_cfs)
    after_tax_em = calc_equity_multiple(levered_atcfs)
    coc_yr1 = calc_cash_on_cash(annual_btcfs[0], equity) if equity > 0 else 0
    coc_yr1_at = calc_cash_on_cash(annual_atcfs[0], equity) if equity > 0 else 0
    dscr_yr1 = calc_dscr(annual_nois[0], annual_ds) if annual_ds > 0 else 0
    yoc = calc_yield_on_cost(annual_nois[0], total_cost)

    # Stabilized metrics (year 3)
    stabilized_noi = annual_nois[min(2, len(annual_nois) - 1)]
    stabilized_yoc = calc_yield_on_cost(stabilized_noi, total_cost)
    stabilized_dscr = calc_dscr(stabilized_noi, annual_ds) if annual_ds > 0 else 0

    metrics = {
        "levered_irr": levered_irr,
        "unlevered_irr": unlevered_irr,
        "after_tax_irr": after_tax_irr,
        "equity_multiple": equity_mult,
        "after_tax_equity_multiple": after_tax_em,
        "cash_on_cash_yr1": coc_yr1,
        "cash_on_cash_yr1_after_tax": coc_yr1_at,
        "dscr_yr1": dscr_yr1,
        "yield_on_cost": yoc,
        "stabilized_yoc": stabilized_yoc,
        "stabilized_dscr": stabilized_dscr,
        "going_in_cap_rate": derived.going_in_cap_rate,
        "exit_cap_rate": exit_cap,
        "price_per_unit": derived.price_per_unit,
        "price_per_sf": derived.price_per_sf,
        "used_variable_growth": has_variable_growth,
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
        "after_tax_cash_flows": [round(cf, 2) for cf in levered_atcfs],
        "annual_nois": [round(n, 2) for n in annual_nois],
        "annual_btcfs": [round(b, 2) for b in annual_btcfs],
        "annual_atcfs": [round(a, 2) for a in annual_atcfs],
    }
