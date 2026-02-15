"""Census Bureau API client for demographic and housing data — city and state level."""

import logging
import requests

logger = logging.getLogger(__name__)

BASE_URL = "https://api.census.gov/data"
TIMEOUT = 15

# State FIPS codes
STATE_FIPS = {
    "AL": "01", "AK": "02", "AZ": "04", "AR": "05", "CA": "06",
    "CO": "08", "CT": "09", "DE": "10", "FL": "12", "GA": "13",
    "HI": "15", "ID": "16", "IL": "17", "IN": "18", "IA": "19",
    "KS": "20", "KY": "21", "LA": "22", "ME": "23", "MD": "24",
    "MA": "25", "MI": "26", "MN": "27", "MS": "28", "MO": "29",
    "MT": "30", "NE": "31", "NV": "32", "NH": "33", "NJ": "34",
    "NM": "35", "NY": "36", "NC": "37", "ND": "38", "OH": "39",
    "OK": "40", "OR": "41", "PA": "42", "RI": "44", "SC": "45",
    "SD": "46", "TN": "47", "TX": "48", "UT": "49", "VT": "50",
    "VA": "51", "WA": "53", "WV": "54", "WI": "55", "WY": "56",
    "DC": "11",
}

# Major city FIPS place codes (Fix #3: city-level data)
CITY_FIPS = {
    # Texas
    ("Austin", "TX"): "05000", ("Houston", "TX"): "35000",
    ("Dallas", "TX"): "19000", ("San Antonio", "TX"): "65000",
    ("Fort Worth", "TX"): "27000", ("El Paso", "TX"): "24000",
    # California
    ("Los Angeles", "CA"): "44000", ("San Francisco", "CA"): "67000",
    ("San Diego", "CA"): "66000", ("San Jose", "CA"): "68000",
    ("Sacramento", "CA"): "64000", ("Oakland", "CA"): "53000",
    # New York
    ("New York", "NY"): "51000", ("Buffalo", "NY"): "11000",
    # Florida
    ("Miami", "FL"): "45000", ("Orlando", "FL"): "53000",
    ("Tampa", "FL"): "71000", ("Jacksonville", "FL"): "35000",
    # Illinois
    ("Chicago", "IL"): "14000",
    # Other major cities
    ("Phoenix", "AZ"): "55000", ("Philadelphia", "PA"): "60000",
    ("Seattle", "WA"): "63000", ("Denver", "CO"): "20000",
    ("Nashville", "TN"): "52006", ("Charlotte", "NC"): "12000",
    ("Atlanta", "GA"): "04000", ("Portland", "OR"): "59000",
    ("Las Vegas", "NV"): "40000", ("Minneapolis", "MN"): "43000",
    ("Detroit", "MI"): "22000", ("Boston", "MA"): "07000",
    ("Baltimore", "MD"): "04000", ("Indianapolis", "IN"): "36003",
    ("Columbus", "OH"): "18000", ("Kansas City", "MO"): "38000",
    ("Raleigh", "NC"): "55000", ("Memphis", "TN"): "48000",
    ("Louisville", "KY"): "48006", ("Milwaukee", "WI"): "53000",
    ("Oklahoma City", "OK"): "55000", ("Tucson", "AZ"): "77000",
    ("New Orleans", "LA"): "55000", ("Cleveland", "OH"): "16000",
    ("Pittsburgh", "PA"): "61000", ("Cincinnati", "OH"): "15000",
    ("St. Louis", "MO"): "65000", ("Salt Lake City", "UT"): "67000",
    ("Washington", "DC"): "50000",
}


class CensusClient:
    """Client for the U.S. Census Bureau API (ACS 5-Year estimates)."""

    def __init__(self):
        self.session = requests.Session()

    def _get_state_fips(self, state_abbr: str) -> str | None:
        return STATE_FIPS.get(state_abbr.upper().strip())

    def _get_city_fips(self, city: str, state: str) -> str | None:
        """Look up FIPS place code for a city."""
        city_clean = city.strip()
        state_clean = state.strip().upper()
        return CITY_FIPS.get((city_clean, state_clean))

    def _fetch_acs_place(self, variables: list[str], city: str, state: str,
                         year: int = 2022) -> dict:
        """Fetch ACS 5-Year data for a specific city/place."""
        state_fips = self._get_state_fips(state)
        place_fips = self._get_city_fips(city, state)
        if not state_fips or not place_fips:
            return {"error": f"No city FIPS for {city}, {state}", "data": {}}

        try:
            var_str = ",".join(["NAME"] + variables)
            url = f"{BASE_URL}/{year}/acs/acs5"
            params = {
                "get": var_str,
                "for": f"place:{place_fips}",
                "in": f"state:{state_fips}",
            }
            resp = self.session.get(url, params=params, timeout=TIMEOUT)
            resp.raise_for_status()
            rows = resp.json()

            if len(rows) < 2:
                return {"error": "No data returned", "data": {}}

            headers = rows[0]
            values = rows[1]
            result = {}
            for h, v in zip(headers, values):
                try:
                    result[h] = float(v) if v and v not in ("null", "N") else None
                except (ValueError, TypeError):
                    result[h] = v

            return {"data": result, "year": year, "level": "city",
                    "city": city, "state": state}

        except Exception as e:
            logger.warning(f"Census city-level API error for {city}, {state}: {e}")
            return {"error": str(e), "data": {}}

    def _fetch_acs_state(self, variables: list[str], state: str,
                         year: int = 2022) -> dict:
        """Fetch ACS 5-Year data for a state (fallback)."""
        state_fips = self._get_state_fips(state)
        if not state_fips:
            return {"error": f"Unknown state: {state}", "data": {}}

        try:
            var_str = ",".join(["NAME"] + variables)
            url = f"{BASE_URL}/{year}/acs/acs5"
            params = {
                "get": var_str,
                "for": f"state:{state_fips}",
            }
            resp = self.session.get(url, params=params, timeout=TIMEOUT)
            resp.raise_for_status()
            rows = resp.json()

            if len(rows) < 2:
                return {"error": "No data returned", "data": {}}

            headers = rows[0]
            values = rows[1]
            result = {}
            for h, v in zip(headers, values):
                try:
                    result[h] = float(v) if v and v not in ("null", "N") else None
                except (ValueError, TypeError):
                    result[h] = v

            return {"data": result, "year": year, "level": "state", "state": state}

        except Exception as e:
            logger.warning(f"Census state-level API error: {e}")
            return {"error": str(e), "data": {}}

    def _fetch_acs(self, variables: list[str], city: str, state: str,
                   year: int = 2022) -> dict:
        """Try city-level first, fall back to state-level."""
        result = self._fetch_acs_place(variables, city, state, year)
        if result["data"]:
            return result
        logger.info(f"City-level data unavailable for {city}, {state}; falling back to state")
        return self._fetch_acs_state(variables, state, year)

    def get_population(self, city: str, state: str) -> dict:
        result = self._fetch_acs(["B01003_001E"], city, state)
        pop = result["data"].get("B01003_001E")
        return {
            "population": int(pop) if pop else None,
            "level": result.get("level", "unknown"),
            "source": f"Census ACS 5-Year ({result.get('level', 'N/A')}-level)",
            "error": result.get("error"),
        }

    def get_median_income(self, city: str, state: str) -> dict:
        result = self._fetch_acs(["B19013_001E"], city, state)
        income = result["data"].get("B19013_001E")
        return {
            "median_income": int(income) if income else None,
            "level": result.get("level", "unknown"),
            "source": f"Census ACS 5-Year ({result.get('level', 'N/A')}-level)",
            "error": result.get("error"),
        }

    def get_median_rent(self, city: str, state: str) -> dict:
        result = self._fetch_acs(["B25064_001E"], city, state)
        rent = result["data"].get("B25064_001E")
        return {
            "median_rent": int(rent) if rent else None,
            "level": result.get("level", "unknown"),
            "source": f"Census ACS 5-Year ({result.get('level', 'N/A')}-level)",
            "error": result.get("error"),
        }

    def get_housing_units(self, city: str, state: str) -> dict:
        variables = ["B25001_001E", "B25002_003E", "B25003_002E", "B25003_003E"]
        result = self._fetch_acs(variables, city, state)
        d = result["data"]

        total = d.get("B25001_001E")
        vacant = d.get("B25002_003E")
        owner = d.get("B25003_002E")
        renter = d.get("B25003_003E")

        vacancy_rate = round((vacant / total) * 100, 1) if total and vacant else None
        renter_pct = round((renter / (owner + renter)) * 100, 1) if owner and renter else None

        return {
            "total_units": int(total) if total else None,
            "vacant_units": int(vacant) if vacant else None,
            "owner_occupied": int(owner) if owner else None,
            "renter_occupied": int(renter) if renter else None,
            "vacancy_rate": vacancy_rate,
            "renter_pct": renter_pct,
            "level": result.get("level", "unknown"),
            "source": f"Census ACS 5-Year ({result.get('level', 'N/A')}-level)",
            "error": result.get("error"),
        }

    def get_all_demographics(self, city: str, state: str) -> dict:
        """Fetch all demographic data — city-level if available, else state."""
        variables = [
            "B01003_001E",   # Population
            "B19013_001E",   # Median Income
            "B25064_001E",   # Median Rent
            "B25001_001E",   # Total Housing Units
            "B25002_003E",   # Vacant Units
            "B25003_002E",   # Owner-Occupied
            "B25003_003E",   # Renter-Occupied
        ]
        result = self._fetch_acs(variables, city, state)
        d = result["data"]

        total_units = d.get("B25001_001E")
        vacant = d.get("B25002_003E")
        owner = d.get("B25003_002E")
        renter = d.get("B25003_003E")
        level = result.get("level", "unknown")

        return {
            "population": int(d["B01003_001E"]) if d.get("B01003_001E") else None,
            "median_income": int(d["B19013_001E"]) if d.get("B19013_001E") else None,
            "median_rent": int(d["B25064_001E"]) if d.get("B25064_001E") else None,
            "total_housing_units": int(total_units) if total_units else None,
            "vacant_units": int(vacant) if vacant else None,
            "vacancy_rate": round((vacant / total_units) * 100, 1) if total_units and vacant else None,
            "renter_pct": round((renter / (owner + renter)) * 100, 1) if owner and renter else None,
            "level": level,
            "city": city,
            "state": state,
            "source": f"Census ACS 5-Year ({level}-level)",
            "error": result.get("error"),
        }
