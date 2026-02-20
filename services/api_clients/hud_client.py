"""HUD Fair Market Rents (FMR) API client.

Data source: HUD USER API — https://www.huduser.gov/hudapi/public/fmr
Coverage: All US metro areas + counties, updated annually.
Year: 2026 FMR (current as of this integration).

FMR represents the 40th percentile gross rents for standard-quality rental units.
Used as a benchmark for affordable/workforce housing underwriting.
"""

import logging
import requests

from config import HUD_API_KEY

logger = logging.getLogger(__name__)

BASE_URL = "https://www.huduser.gov/hudapi/public"
TIMEOUT = 15

# State abbreviation → full state code for API
STATE_CODES = {
    "AL": "AL", "AK": "AK", "AZ": "AZ", "AR": "AR", "CA": "CA",
    "CO": "CO", "CT": "CT", "DE": "DE", "FL": "FL", "GA": "GA",
    "HI": "HI", "ID": "ID", "IL": "IL", "IN": "IN", "IA": "IA",
    "KS": "KS", "KY": "KY", "LA": "LA", "ME": "ME", "MD": "MD",
    "MA": "MA", "MI": "MI", "MN": "MN", "MS": "MS", "MO": "MO",
    "MT": "MT", "NE": "NE", "NV": "NV", "NH": "NH", "NJ": "NJ",
    "NM": "NM", "NY": "NY", "NC": "NC", "ND": "ND", "OH": "OH",
    "OK": "OK", "OR": "OR", "PA": "PA", "RI": "RI", "SC": "SC",
    "SD": "SD", "TN": "TN", "TX": "TX", "UT": "UT", "VT": "VT",
    "VA": "VA", "WA": "WA", "WV": "WV", "WI": "WI", "WY": "WY",
    "DC": "DC",
}


class HUDClient:
    """Client for HUD Fair Market Rents API."""

    def __init__(self, api_key: str = None):
        self.api_key = api_key or HUD_API_KEY
        self.session = requests.Session()
        if self.api_key:
            self.session.headers.update({
                "Authorization": f"Bearer {self.api_key}",
                "Accept": "application/json",
            })

    def _is_configured(self) -> bool:
        return bool(self.api_key)

    def _get(self, endpoint: str, params: dict = None) -> dict | list | None:
        if not self._is_configured():
            logger.warning("HUD_API_KEY not configured")
            return None
        try:
            url = f"{BASE_URL}/{endpoint.lstrip('/')}"
            resp = self.session.get(url, params=params, timeout=TIMEOUT)
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            logger.warning(f"HUD API error ({endpoint}): {e}")
            return None

    def _normalize(self, text: str) -> str:
        return text.lower().strip()

    def get_fmr_by_city(self, city: str, state: str) -> dict:
        """Look up Fair Market Rents for a city/metro area.

        Fetches all metro areas for the state, then fuzzy-matches the city
        name to the MSA name. Returns FMR by bedroom count (Efficiency through 4BR).

        Args:
            city: City name (e.g. "Austin")
            state: State abbreviation (e.g. "TX")

        Returns:
            Dict with fmr_by_bedroom, metro_name, year, percentile, and source.
        """
        state_upper = state.upper()
        if state_upper not in STATE_CODES:
            return {"available": False, "note": f"Unknown state: {state}"}

        data = self._get(f"/fmr/statedata/{state_upper}")
        if not data or "data" not in data:
            return {"available": False, "note": "HUD API returned no data"}

        metro_areas = data["data"].get("metroareas", [])
        if not metro_areas:
            return {"available": False, "note": f"No metro areas found for {state_upper}"}

        city_norm = self._normalize(city)
        best_match = None
        best_score = 0

        for metro in metro_areas:
            metro_name = metro.get("metro_name", "")
            metro_norm = self._normalize(metro_name)

            if city_norm not in metro_norm:
                continue

            # Scoring rules (higher = better):
            # 1. Metro name STARTS with city → highest priority (e.g. "Austin-Round Rock" for "Austin")
            # 2. Match is a whole word boundary → better than substring
            # 3. Penalize "County" entries when city match is at the start of a real MSA
            starts_with_city = metro_norm.startswith(city_norm)
            is_county = "county" in metro_norm

            score = len(city_norm) / max(len(metro_norm), 1)
            if starts_with_city:
                score += 1.0  # Strong boost for starts-with
            if is_county and not city_norm.endswith("county"):
                score -= 0.5  # Penalize county entries when city itself isn't a county

            if score > best_score:
                best_score = score
                best_match = metro

        if not best_match:
            return {
                "available": False,
                "city": city,
                "state": state,
                "note": (
                    f"No HUD metro area found matching '{city}' in {state}. "
                    f"HUD FMR data covers MSAs — small cities may not have a dedicated entry."
                ),
            }

        fmr_by_br = {
            "efficiency": best_match.get("Efficiency"),
            "one_br": best_match.get("One-Bedroom"),
            "two_br": best_match.get("Two-Bedroom"),
            "three_br": best_match.get("Three-Bedroom"),
            "four_br": best_match.get("Four-Bedroom"),
        }

        return {
            "available": True,
            "city": city,
            "state": state,
            "metro_name": best_match.get("metro_name", ""),
            "entity_code": best_match.get("code", ""),
            "year": best_match.get("year", "2026"),
            "fmr_percentile": best_match.get("FMR Percentile", 40),
            "fmr_by_bedroom": fmr_by_br,
            "source": "HUD FMR API 2026",
            "note": (
                f"HUD Fair Market Rents for {best_match.get('metro_name', '')}. "
                f"Represents the 40th percentile gross rent for standard-quality units."
            ),
        }

    def get_fmr_for_avg_bedrooms(self, city: str, state: str,
                                  avg_bedrooms: int = 2) -> dict:
        """Get the single FMR value most relevant to the property's avg bedroom count.

        Args:
            city: City name
            state: State abbreviation
            avg_bedrooms: Average bedroom count per unit (0=efficiency, 1-4=BR)

        Returns:
            Dict including the matched FMR rent and all bedroom FMRs.
        """
        result = self.get_fmr_by_city(city, state)
        if not result.get("available"):
            return result

        br_map = {
            0: "efficiency",
            1: "one_br",
            2: "two_br",
            3: "three_br",
            4: "four_br",
        }
        br_key = br_map.get(min(avg_bedrooms, 4), "two_br")
        matched_fmr = result["fmr_by_bedroom"].get(br_key)

        result["matched_bedroom_key"] = br_key
        result["matched_fmr"] = matched_fmr
        result["avg_bedrooms_requested"] = avg_bedrooms
        return result
