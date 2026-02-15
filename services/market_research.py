"""Market research using FRED, Census, and BLS APIs with graceful fallbacks."""

import logging

from services.api_clients.fred_client import FREDClient
from services.api_clients.census_client import CensusClient
from services.api_clients.bls_client import BLSClient

logger = logging.getLogger(__name__)

_cache: dict[str, dict] = {}

CAP_RATE_SPREADS = {
    "Multifamily": 1.75,
    "Office": 2.50,
    "Retail": 2.25,
    "Industrial": 2.00,
}

FALLBACK_CAP_RATES = {
    "Multifamily": 5.25,
    "Office": 6.50,
    "Retail": 6.25,
    "Industrial": 5.75,
}

fred = FREDClient()
census = CensusClient()
bls = BLSClient()


def search_comps(property_type: str, city: str, state: str) -> dict:
    """Generate comparable sales calibrated to actual city-level market data.

    Fix #4: Comps are synthetic but explicitly labeled and better calibrated
    using city-level Census data when available.
    """
    cache_key = f"comps_{property_type}_{city}_{state}"
    if cache_key in _cache:
        return _cache[cache_key]

    # Use city-level Census data for calibration
    census_data = census.get_all_demographics(city, state)
    median_rent = census_data.get("median_rent")
    median_income = census_data.get("median_income")
    data_level = census_data.get("level", "unknown")

    base_ppu = _estimate_price_per_unit(property_type, median_rent, median_income)
    base_cap = FALLBACK_CAP_RATES.get(property_type, 6.0)

    import random
    random.seed(hash(f"{city}{state}{property_type}"))

    comps = []
    for i in range(5):
        variation = random.uniform(-0.12, 0.12)
        cap_var = random.uniform(-0.5, 0.5)
        units = random.choice([60, 80, 95, 110, 120, 150, 175, 200])
        comps.append({
            "name": f"{city} {property_type} Comp {i+1}",
            "price_per_unit": int(base_ppu * (1 + variation)),
            "cap_rate": round(base_cap + cap_var, 2),
            "year": random.choice([2023, 2024, 2024, 2025]),
            "units": units,
            "synthetic": True,  # Fix #4: Explicit labeling
        })

    data = {
        "comps": comps,
        "source": "synthetic_calibrated",
        "data_level": data_level,
        "note": (
            f"Comps are SYNTHETIC estimates calibrated to {data_level}-level "
            f"Census data (median rent: ${median_rent:,}/mo, median income: "
            f"${median_income:,}). They are NOT real transaction records. "
            f"Use actual broker comps for investment decisions."
            if median_rent and median_income else
            "Comps are synthetic defaults. No Census data available for calibration."
        ),
        "calibration_data": {
            "median_rent": median_rent,
            "median_income": median_income,
            "data_level": data_level,
        },
    }
    _cache[cache_key] = data
    return data


def search_cap_rates(property_type: str, city: str) -> dict:
    """Derive market cap rates from FRED Treasury data + property-type spread."""
    cache_key = f"cap_rates_{property_type}_{city}"
    if cache_key in _cache:
        return _cache[cache_key]

    treasury = fred.get_treasury_rates()
    mortgage = fred.get_mortgage_rates()
    t10 = treasury.get("treasury_10yr")

    spread = CAP_RATE_SPREADS.get(property_type, 2.25)

    if t10 is not None:
        derived_cap = round(t10 + spread, 2)
        market_caps = [
            round(derived_cap - 0.50, 2),
            round(derived_cap - 0.25, 2),
            derived_cap,
            round(derived_cap + 0.25, 2),
            round(derived_cap + 0.50, 2),
        ]
        source = "FRED_derived"
    else:
        derived_cap = FALLBACK_CAP_RATES.get(property_type, 6.0)
        market_caps = [derived_cap]
        source = "defaults"

    data = {
        "market_cap_rates": market_caps,
        "average_cap_rate": derived_cap,
        "treasury_10yr": t10,
        "treasury_2yr": treasury.get("treasury_2yr"),
        "spread_used": spread,
        "mortgage_30yr": mortgage.get("current_rate"),
        "search_results": [{
            "source": "FRED 10-Year Treasury + Property Spread",
            "snippet": f"10Y Treasury: {t10}% + {spread}% spread = {derived_cap}% cap rate"
                       if t10 else "Using fallback cap rate defaults",
        }],
        "source": source,
    }
    _cache[cache_key] = data
    return data


def search_demographics(city: str, state: str) -> dict:
    """Gather demographic data from Census (city-level) and BLS."""
    cache_key = f"demo_{city}_{state}"
    if cache_key in _cache:
        return _cache[cache_key]

    # Fix #3: Try city-level first, fall back to state
    census_data = census.get_all_demographics(city, state)
    labor_data = bls.get_all_labor_data()
    fred_unemployment = fred.get_unemployment_rate()

    pop = census_data.get("population")
    income = census_data.get("median_income")
    median_rent = census_data.get("median_rent")
    vacancy = census_data.get("vacancy_rate")
    renter_pct = census_data.get("renter_pct")
    data_level = census_data.get("level", "unknown")
    unemp = (fred_unemployment.get("current_rate")
             or labor_data.get("unemployment", {}).get("rate"))

    has_data = pop is not None or income is not None
    summary = _build_api_summary(city, state, data_level, pop, income,
                                 median_rent, unemp, vacancy, renter_pct)

    search_results = []
    level_label = f"({data_level}-level)" if data_level != "unknown" else ""
    if pop:
        search_results.append({
            "source": f"Census Bureau ACS {level_label}",
            "snippet": f"{city} population: {pop:,}; Median income: ${income:,}" if income else f"{city} population: {pop:,}",
        })
    if unemp:
        search_results.append({
            "source": "FRED / BLS",
            "snippet": f"National unemployment rate: {unemp}%",
        })
    if median_rent:
        search_results.append({
            "source": f"Census Bureau ACS {level_label}",
            "snippet": f"{city} median gross rent: ${median_rent:,}",
        })

    data = {
        "search_results": search_results,
        "city": city,
        "state": state,
        "source": f"FRED_Census_BLS ({data_level})" if has_data else "defaults",
        "summary": summary,
        "structured": {
            "population": pop,
            "median_income": income,
            "median_rent": median_rent,
            "unemployment_rate": unemp,
            "vacancy_rate": vacancy,
            "renter_pct": renter_pct,
            "employment_thousands": labor_data.get("employment", {}).get("total_thousands"),
            "data_level": data_level,
        },
    }
    _cache[cache_key] = data
    return data


def search_rent_trends(property_type: str, city: str) -> dict:
    """Get rent growth trends from FRED CPI Shelter + Census median rent."""
    cache_key = f"rent_{property_type}_{city}"
    if cache_key in _cache:
        return _cache[cache_key]

    cpi_shelter = fred.get_cpi_shelter(limit=36)
    annual_rates = cpi_shelter.get("annual_growth_rates", [])

    growth_rates = [r["growth_rate"] for r in annual_rates[-5:]] if annual_rates else []
    avg_growth = round(sum(growth_rates) / len(growth_rates), 2) if growth_rates else 3.0

    search_results = []
    if growth_rates:
        search_results.append({
            "source": "FRED CPI Shelter",
            "snippet": f"Recent shelter inflation rates: {', '.join(f'{r:.1f}%' for r in growth_rates[-3:])}",
        })

    data = {
        "rent_growth_rates": growth_rates,
        "average_growth": avg_growth,
        "search_results": search_results,
        "source": "FRED_CPI_Shelter" if growth_rates else "defaults",
        "cpi_shelter_data": annual_rates[-12:] if annual_rates else [],
    }
    _cache[cache_key] = data
    return data


def run_full_research(property_type: str, city: str, state: str) -> dict:
    """Run all market research and return combined results."""
    return {
        "comps": search_comps(property_type, city, state),
        "cap_rates": search_cap_rates(property_type, city),
        "demographics": search_demographics(city, state),
        "rent_trends": search_rent_trends(property_type, city),
    }


def _estimate_price_per_unit(property_type: str, median_rent: int | None,
                             median_income: int | None) -> int:
    if median_rent and median_rent > 0:
        annual_rent = median_rent * 12
        grm = {"Multifamily": 13, "Office": 11, "Retail": 10, "Industrial": 12}
        multiplier = grm.get(property_type, 12)
        return int(annual_rent * multiplier)
    defaults = {"Multifamily": 155000, "Office": 180000, "Retail": 170000, "Industrial": 140000}
    return defaults.get(property_type, 155000)


def _build_api_summary(city: str, state: str, data_level: str,
                       population: int | None, median_income: int | None,
                       median_rent: int | None, unemployment: float | None,
                       vacancy: float | None, renter_pct: float | None) -> str:
    level_note = f" ({data_level}-level data)" if data_level != "unknown" else ""
    parts = [f"{city}, {state}{level_note}"]

    if population:
        if population >= 1_000_000:
            parts.append(f"has a population of {population/1e6:.1f} million")
        else:
            parts.append(f"has a population of {population:,}")

    if median_income:
        parts.append(f"median household income of ${median_income:,}")
    if median_rent:
        parts.append(f"median gross rent of ${median_rent:,}/month")
    if unemployment:
        parts.append(f"a national unemployment rate of {unemployment}%")
    if vacancy:
        parts.append(f"a housing vacancy rate of {vacancy}%")
    if renter_pct:
        parts.append(f"({renter_pct}% renter-occupied)")

    summary = " with ".join(parts[:2])
    if len(parts) > 2:
        summary += ", " + ", ".join(parts[2:])
    summary += ". "

    if median_income and median_income > 55000:
        summary += "The area demonstrates above-average income levels, supporting demand for quality rental housing. "
    if unemployment and unemployment < 5.0:
        summary += "Low unemployment indicates a healthy local economy. "

    summary += "Data sourced from FRED, Census Bureau ACS, and BLS."
    return summary
