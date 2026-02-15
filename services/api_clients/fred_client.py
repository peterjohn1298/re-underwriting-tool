"""FRED API client for macroeconomic and real estate data."""

import logging
from datetime import datetime, timedelta
import requests

from config import FRED_API_KEY

logger = logging.getLogger(__name__)

BASE_URL = "https://api.stlouisfed.org/fred"
TIMEOUT = 15

# FRED Series IDs
SERIES = {
    "mortgage_30yr": "MORTGAGE30US",       # 30-Year Fixed Rate Mortgage
    "treasury_10yr": "DGS10",              # 10-Year Treasury Constant Maturity
    "treasury_2yr": "DGS2",                # 2-Year Treasury
    "cpi_all": "CPIAUCSL",                 # CPI for All Urban Consumers
    "cpi_shelter": "CUSR0000SAH1",         # CPI: Shelter
    "unemployment": "UNRATE",              # Civilian Unemployment Rate
    "housing_starts": "HOUST",             # Housing Starts
    "rental_vacancy": "RRVRUSQ156N",       # Rental Vacancy Rate
    "fed_funds": "FEDFUNDS",               # Federal Funds Effective Rate
    "real_gdp": "A191RL1Q225SBEA",         # Real GDP Growth Rate
}


class FREDClient:
    """Client for the Federal Reserve Economic Data (FRED) API."""

    def __init__(self, api_key: str = None):
        self.api_key = api_key or FRED_API_KEY
        self.session = requests.Session()

    def _fetch_series(self, series_id: str, limit: int = 12,
                      sort_order: str = "desc") -> dict:
        """Fetch observations for a FRED series."""
        if not self.api_key or self.api_key == "your_fred_api_key_here":
            return {"error": "FRED_API_KEY not configured", "data": []}

        try:
            url = f"{BASE_URL}/series/observations"
            params = {
                "series_id": series_id,
                "api_key": self.api_key,
                "file_type": "json",
                "sort_order": sort_order,
                "limit": limit,
            }
            resp = self.session.get(url, params=params, timeout=TIMEOUT)
            resp.raise_for_status()
            data = resp.json()

            observations = []
            for obs in data.get("observations", []):
                if obs.get("value") and obs["value"] != ".":
                    observations.append({
                        "date": obs["date"],
                        "value": float(obs["value"]),
                    })
            return {"data": observations, "series_id": series_id}

        except Exception as e:
            logger.warning(f"FRED API error for {series_id}: {e}")
            return {"error": str(e), "data": [], "series_id": series_id}

    def _latest_value(self, series_id: str) -> float | None:
        """Get the most recent value for a series."""
        result = self._fetch_series(series_id, limit=1)
        if result["data"]:
            return result["data"][0]["value"]
        return None

    def get_mortgage_rates(self) -> dict:
        """Get current 30-year fixed mortgage rate."""
        result = self._fetch_series(SERIES["mortgage_30yr"], limit=12)
        latest = result["data"][0]["value"] if result["data"] else None
        return {
            "current_rate": latest,
            "recent_data": result["data"][:6],
            "source": "FRED (MORTGAGE30US)",
            "error": result.get("error"),
        }

    def get_treasury_rates(self) -> dict:
        """Get 10-Year and 2-Year Treasury rates."""
        t10 = self._latest_value(SERIES["treasury_10yr"])
        t2 = self._latest_value(SERIES["treasury_2yr"])
        return {
            "treasury_10yr": t10,
            "treasury_2yr": t2,
            "spread": round(t10 - t2, 2) if t10 and t2 else None,
            "source": "FRED (DGS10, DGS2)",
        }

    def get_cpi_data(self) -> dict:
        """Get CPI inflation data."""
        result = self._fetch_series(SERIES["cpi_all"], limit=24)
        data = result["data"]
        yoy_change = None
        if len(data) >= 13:
            current = data[0]["value"]
            year_ago = data[12]["value"]
            yoy_change = round(((current - year_ago) / year_ago) * 100, 2)

        return {
            "latest_cpi": data[0]["value"] if data else None,
            "yoy_inflation": yoy_change,
            "recent_data": data[:6],
            "source": "FRED (CPIAUCSL)",
            "error": result.get("error"),
        }

    def get_cpi_shelter(self, limit: int = 60) -> dict:
        """Get CPI Shelter component — used for rent growth modeling."""
        result = self._fetch_series(SERIES["cpi_shelter"], limit=limit)
        data = result["data"]

        # Calculate annual growth rates from monthly data
        annual_rates = []
        if len(data) >= 13:
            # data is in desc order, reverse for chronological
            chronological = list(reversed(data))
            for i in range(12, len(chronological)):
                current = chronological[i]["value"]
                year_ago = chronological[i - 12]["value"]
                if year_ago > 0:
                    rate = ((current - year_ago) / year_ago) * 100
                    annual_rates.append({
                        "date": chronological[i]["date"],
                        "growth_rate": round(rate, 2),
                    })

        return {
            "latest_value": data[0]["value"] if data else None,
            "annual_growth_rates": annual_rates,
            "raw_data": data,
            "source": "FRED (CUSR0000SAH1)",
            "error": result.get("error"),
        }

    def get_unemployment_rate(self) -> dict:
        """Get current unemployment rate."""
        result = self._fetch_series(SERIES["unemployment"], limit=12)
        return {
            "current_rate": result["data"][0]["value"] if result["data"] else None,
            "recent_data": result["data"][:6],
            "source": "FRED (UNRATE)",
            "error": result.get("error"),
        }

    def get_housing_starts(self) -> dict:
        """Get housing starts data (thousands of units)."""
        result = self._fetch_series(SERIES["housing_starts"], limit=12)
        return {
            "current": result["data"][0]["value"] if result["data"] else None,
            "recent_data": result["data"][:6],
            "source": "FRED (HOUST)",
            "error": result.get("error"),
        }

    def get_rental_vacancy_rate(self) -> dict:
        """Get rental vacancy rate (quarterly)."""
        result = self._fetch_series(SERIES["rental_vacancy"], limit=8)
        return {
            "current_rate": result["data"][0]["value"] if result["data"] else None,
            "recent_data": result["data"][:4],
            "source": "FRED (RRVRUSQ156N)",
            "error": result.get("error"),
        }

    def get_all_macro_data(self) -> dict:
        """Fetch all macro indicators in one call."""
        return {
            "mortgage_rates": self.get_mortgage_rates(),
            "treasury_rates": self.get_treasury_rates(),
            "cpi": self.get_cpi_data(),
            "cpi_shelter": self.get_cpi_shelter(),
            "unemployment": self.get_unemployment_rate(),
            "housing_starts": self.get_housing_starts(),
            "rental_vacancy": self.get_rental_vacancy_rate(),
        }
