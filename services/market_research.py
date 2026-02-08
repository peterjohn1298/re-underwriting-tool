"""Market research via web scraping with graceful fallbacks."""

import re
import logging
import requests
from bs4 import BeautifulSoup

logger = logging.getLogger(__name__)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}
TIMEOUT = 10

# In-memory cache for session
_cache: dict[str, dict] = {}


def _google_search(query: str, num_results: int = 5) -> list[dict]:
    """Perform a Google search and return titles + snippets."""
    try:
        url = "https://www.google.com/search"
        params = {"q": query, "num": num_results}
        resp = requests.get(url, params=params, headers=HEADERS, timeout=TIMEOUT)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        results = []
        for div in soup.select("div.g, div[data-hveid]"):
            title_el = div.select_one("h3")
            snippet_el = div.select_one("div.VwiC3b, span.aCOpRe, div[data-sncf]")
            if title_el:
                results.append({
                    "title": title_el.get_text(strip=True),
                    "snippet": snippet_el.get_text(strip=True) if snippet_el else "",
                })
        return results[:num_results]
    except Exception as e:
        logger.warning(f"Google search failed for '{query}': {e}")
        return []


def _extract_numbers(text: str) -> list[float]:
    """Extract numbers (including decimals and percentages) from text."""
    matches = re.findall(r'[\d,]+\.?\d*%?', text)
    nums = []
    for m in matches:
        try:
            clean = m.replace(",", "").replace("%", "")
            nums.append(float(clean))
        except ValueError:
            pass
    return nums


def search_comps(property_type: str, city: str, state: str) -> dict:
    """Search for comparable sales data."""
    cache_key = f"comps_{property_type}_{city}_{state}"
    if cache_key in _cache:
        return _cache[cache_key]

    query = f"{property_type} commercial real estate recent sales {city} {state} price per unit cap rate 2024 2025"
    results = _google_search(query)

    comps = []
    for r in results:
        snippet = r["snippet"]
        numbers = _extract_numbers(snippet)
        comps.append({
            "source": r["title"],
            "details": snippet,
            "extracted_numbers": numbers,
        })

    data = {
        "comps": comps if comps else _default_comps(property_type, city),
        "source": "web_search" if comps else "defaults",
    }
    _cache[cache_key] = data
    return data


def search_cap_rates(property_type: str, city: str) -> dict:
    """Search for market cap rates."""
    cache_key = f"cap_rates_{property_type}_{city}"
    if cache_key in _cache:
        return _cache[cache_key]

    query = f"{property_type} cap rate {city} market 2024 2025 average"
    results = _google_search(query)

    cap_rates = []
    for r in results:
        numbers = _extract_numbers(r["snippet"])
        for n in numbers:
            if 3.0 <= n <= 12.0:
                cap_rates.append(n)

    avg_cap = sum(cap_rates) / len(cap_rates) if cap_rates else None

    data = {
        "market_cap_rates": cap_rates[:5],
        "average_cap_rate": round(avg_cap, 2) if avg_cap else _default_cap_rate(property_type),
        "search_results": [{"source": r["title"], "snippet": r["snippet"]} for r in results],
        "source": "web_search" if cap_rates else "defaults",
    }
    _cache[cache_key] = data
    return data


def search_demographics(city: str, state: str) -> dict:
    """Search for demographic data."""
    cache_key = f"demo_{city}_{state}"
    if cache_key in _cache:
        return _cache[cache_key]

    query = f"{city} {state} population employment rate median household income growth 2024"
    results = _google_search(query)

    data = {
        "search_results": [{"source": r["title"], "snippet": r["snippet"]} for r in results],
        "city": city,
        "state": state,
        "source": "web_search" if results else "defaults",
        "summary": _build_demo_summary(results, city, state),
    }
    _cache[cache_key] = data
    return data


def search_rent_trends(property_type: str, city: str) -> dict:
    """Search for rent growth trends."""
    cache_key = f"rent_{property_type}_{city}"
    if cache_key in _cache:
        return _cache[cache_key]

    query = f"{property_type} rent growth trends {city} 2024 2025 year over year"
    results = _google_search(query)

    growth_rates = []
    for r in results:
        numbers = _extract_numbers(r["snippet"])
        for n in numbers:
            if -5.0 <= n <= 15.0:
                growth_rates.append(n)

    data = {
        "rent_growth_rates": growth_rates[:5],
        "average_growth": round(sum(growth_rates) / len(growth_rates), 2) if growth_rates else 3.0,
        "search_results": [{"source": r["title"], "snippet": r["snippet"]} for r in results],
        "source": "web_search" if growth_rates else "defaults",
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


def _default_comps(property_type: str, city: str) -> list[dict]:
    """Fallback comp data when web search fails."""
    base_comps = [
        {"name": f"{city} {property_type} Comp 1", "price_per_unit": 150000, "cap_rate": 5.5, "year": 2024, "units": 120},
        {"name": f"{city} {property_type} Comp 2", "price_per_unit": 165000, "cap_rate": 5.25, "year": 2024, "units": 95},
        {"name": f"{city} {property_type} Comp 3", "price_per_unit": 140000, "cap_rate": 5.75, "year": 2023, "units": 150},
        {"name": f"{city} {property_type} Comp 4", "price_per_unit": 175000, "cap_rate": 5.0, "year": 2024, "units": 80},
        {"name": f"{city} {property_type} Comp 5", "price_per_unit": 155000, "cap_rate": 5.5, "year": 2023, "units": 110},
    ]
    return base_comps


def _default_cap_rate(property_type: str) -> float:
    """Fallback cap rates by property type."""
    defaults = {
        "Multifamily": 5.25,
        "Office": 6.50,
        "Retail": 6.25,
        "Industrial": 5.75,
    }
    return defaults.get(property_type, 6.0)


def _build_demo_summary(results: list[dict], city: str, state: str) -> str:
    """Build a narrative summary from search results."""
    if not results:
        return (
            f"{city}, {state} is a growing market with positive economic indicators. "
            f"The metropolitan area has seen steady population and employment growth, "
            f"supporting demand for commercial real estate."
        )
    snippets = " ".join(r["snippet"] for r in results[:3])
    if len(snippets) > 500:
        snippets = snippets[:500] + "..."
    return f"Based on market research: {snippets}"
