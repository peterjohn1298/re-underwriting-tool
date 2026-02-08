from dataclasses import dataclass, field


@dataclass
class DealInputs:
    """Raw deal sheet inputs — exactly what the user types in."""
    property_type: str = "Multifamily - Class B"
    address: str = ""
    year_built: int = 2000
    purchase_price: float = 0.0
    current_noi: float = 0.0
    total_units: int = 0
    total_sf: float = 0.0
    in_place_rent: float = 0.0      # $/unit/month
    market_rent: float = 0.0        # $/unit/month
    occupancy: float = 0.92         # 0-1
    deferred_maintenance: float = 0.0
    planned_capex: float = 0.0
    capex_description: str = ""
    hold_period_years: int = 7
    # Optional overrides (defaults applied if not set)
    ltv: float = 0.65
    interest_rate: float = 0.0675
    amortization_years: int = 30
    loan_term_years: int = 10
    io_period_years: int = 0
    closing_costs_pct: float = 0.03
    revenue_growth_rate: float = 0.03
    expense_growth_rate: float = 0.03
    management_fee_pct: float = 0.035
    exit_cap_rate_spread: float = 0.0025  # +25bps
    sale_costs_pct: float = 0.025
    replacement_reserves_per_unit: float = 250.0


@dataclass
class DerivedAssumptions:
    """Everything derived from the raw deal inputs."""
    # Parsed from address
    property_name: str = ""
    city: str = ""
    state: str = ""
    zip_code: str = ""
    street_address: str = ""
    # Parsed from property_type
    asset_type: str = "Multifamily"  # base type
    asset_class: str = "Class B"

    # Revenue
    gross_potential_rent: float = 0.0    # in-place rent × units × 12
    market_gpr: float = 0.0             # market rent × units × 12
    vacancy_rate: float = 0.08
    effective_gross_income: float = 0.0
    other_income: float = 0.0
    rent_premium_potential: float = 0.0  # market vs in-place spread

    # Expenses (backed out from NOI)
    total_operating_expenses: float = 0.0
    expense_ratio: float = 0.0
    property_tax: float = 0.0
    insurance: float = 0.0
    utilities: float = 0.0
    repairs_maintenance: float = 0.0
    general_admin: float = 0.0
    other_expenses: float = 0.0

    # Capital
    total_capex: float = 0.0            # deferred + planned
    capex_per_unit: float = 0.0
    total_project_cost: float = 0.0
    loan_amount: float = 0.0
    equity_required: float = 0.0

    # Rates
    going_in_cap_rate: float = 0.0
    exit_cap_rate: float = 0.0
    price_per_unit: float = 0.0
    price_per_sf: float = 0.0


def parse_address(address: str) -> dict:
    """Parse a full address string into components."""
    parts = [p.strip() for p in address.split(",")]
    result = {"street": "", "city": "", "state": "", "zip": ""}

    if len(parts) >= 1:
        result["street"] = parts[0]
    if len(parts) >= 2:
        result["city"] = parts[1]
    if len(parts) >= 3:
        # "TX 78751" or "TX"
        state_zip = parts[2].strip().split()
        result["state"] = state_zip[0] if state_zip else ""
        result["zip"] = state_zip[1] if len(state_zip) > 1 else ""
    if len(parts) >= 4:
        result["zip"] = parts[3].strip()

    return result


def parse_property_type(prop_type: str) -> tuple[str, str]:
    """Parse 'Multifamily - Class B' into ('Multifamily', 'Class B')."""
    if " - " in prop_type:
        parts = prop_type.split(" - ", 1)
        return parts[0].strip(), parts[1].strip()
    return prop_type.strip(), ""


def derive_assumptions(deal: DealInputs) -> DerivedAssumptions:
    """Derive all underwriting assumptions from raw deal inputs."""
    d = DerivedAssumptions()

    # Parse address
    addr = parse_address(deal.address)
    d.street_address = addr["street"]
    d.city = addr["city"]
    d.state = addr["state"]
    d.zip_code = addr["zip"]
    d.property_name = f"{addr['street']}" if addr["street"] else "Subject Property"

    # Parse type
    d.asset_type, d.asset_class = parse_property_type(deal.property_type)

    # --- Revenue ---
    d.gross_potential_rent = deal.in_place_rent * deal.total_units * 12
    d.market_gpr = deal.market_rent * deal.total_units * 12 if deal.market_rent > 0 else d.gross_potential_rent
    d.vacancy_rate = 1.0 - deal.occupancy
    d.effective_gross_income = d.gross_potential_rent * deal.occupancy
    d.rent_premium_potential = deal.market_rent - deal.in_place_rent if deal.market_rent > 0 else 0

    # Other income estimate (laundry, parking, etc.) — ~5% of GPR for multifamily
    if "Multifamily" in d.asset_type:
        d.other_income = d.gross_potential_rent * 0.05
    else:
        d.other_income = d.gross_potential_rent * 0.03
    d.effective_gross_income += d.other_income

    # --- Expenses (back into from NOI) ---
    # Current NOI = EGI - Total Expenses
    # Total Expenses = EGI - NOI
    d.total_operating_expenses = d.effective_gross_income - deal.current_noi
    if d.total_operating_expenses < 0:
        d.total_operating_expenses = 0

    if d.effective_gross_income > 0:
        d.expense_ratio = d.total_operating_expenses / d.effective_gross_income
    else:
        d.expense_ratio = 0.45  # default

    # Break down expenses into categories (industry-typical allocation)
    total_ex = d.total_operating_expenses
    mgmt = d.effective_gross_income * deal.management_fee_pct
    remaining = total_ex - mgmt
    if remaining < 0:
        remaining = total_ex
        mgmt = 0

    d.property_tax = remaining * 0.35
    d.insurance = remaining * 0.15
    d.utilities = remaining * 0.15
    d.repairs_maintenance = remaining * 0.15
    d.general_admin = remaining * 0.10
    d.other_expenses = remaining * 0.10

    # --- Capital ---
    d.total_capex = deal.deferred_maintenance + deal.planned_capex
    d.capex_per_unit = d.total_capex / deal.total_units if deal.total_units > 0 else 0

    closing = deal.purchase_price * deal.closing_costs_pct
    d.total_project_cost = deal.purchase_price + closing + d.total_capex
    d.loan_amount = deal.purchase_price * deal.ltv
    d.equity_required = d.total_project_cost - d.loan_amount

    # --- Rates ---
    if deal.purchase_price > 0:
        d.going_in_cap_rate = deal.current_noi / deal.purchase_price
        d.price_per_unit = deal.purchase_price / deal.total_units if deal.total_units > 0 else 0
        d.price_per_sf = deal.purchase_price / deal.total_sf if deal.total_sf > 0 else 0
    d.exit_cap_rate = d.going_in_cap_rate + deal.exit_cap_rate_spread

    return d


def to_full_dict(deal: DealInputs, derived: DerivedAssumptions) -> dict:
    """Convert to dict for templates and generators."""
    return {
        "deal": {
            "property_type": deal.property_type,
            "address": deal.address,
            "year_built": deal.year_built,
            "purchase_price": deal.purchase_price,
            "current_noi": deal.current_noi,
            "total_units": deal.total_units,
            "total_sf": deal.total_sf,
            "in_place_rent": deal.in_place_rent,
            "market_rent": deal.market_rent,
            "occupancy": deal.occupancy,
            "deferred_maintenance": deal.deferred_maintenance,
            "planned_capex": deal.planned_capex,
            "capex_description": deal.capex_description,
            "hold_period_years": deal.hold_period_years,
            "ltv": deal.ltv,
            "interest_rate": deal.interest_rate,
            "amortization_years": deal.amortization_years,
            "loan_term_years": deal.loan_term_years,
            "io_period_years": deal.io_period_years,
            "closing_costs_pct": deal.closing_costs_pct,
            "revenue_growth_rate": deal.revenue_growth_rate,
            "expense_growth_rate": deal.expense_growth_rate,
            "management_fee_pct": deal.management_fee_pct,
            "exit_cap_rate_spread": deal.exit_cap_rate_spread,
            "sale_costs_pct": deal.sale_costs_pct,
            "replacement_reserves_per_unit": deal.replacement_reserves_per_unit,
        },
        "derived": {
            "property_name": derived.property_name,
            "city": derived.city,
            "state": derived.state,
            "zip_code": derived.zip_code,
            "street_address": derived.street_address,
            "asset_type": derived.asset_type,
            "asset_class": derived.asset_class,
            "gross_potential_rent": derived.gross_potential_rent,
            "market_gpr": derived.market_gpr,
            "vacancy_rate": derived.vacancy_rate,
            "effective_gross_income": derived.effective_gross_income,
            "other_income": derived.other_income,
            "rent_premium_potential": derived.rent_premium_potential,
            "total_operating_expenses": derived.total_operating_expenses,
            "expense_ratio": derived.expense_ratio,
            "property_tax": derived.property_tax,
            "insurance": derived.insurance,
            "utilities": derived.utilities,
            "repairs_maintenance": derived.repairs_maintenance,
            "general_admin": derived.general_admin,
            "other_expenses": derived.other_expenses,
            "total_capex": derived.total_capex,
            "capex_per_unit": derived.capex_per_unit,
            "total_project_cost": derived.total_project_cost,
            "loan_amount": derived.loan_amount,
            "equity_required": derived.equity_required,
            "going_in_cap_rate": derived.going_in_cap_rate,
            "exit_cap_rate": derived.exit_cap_rate,
            "price_per_unit": derived.price_per_unit,
            "price_per_sf": derived.price_per_sf,
        },
    }
