"""Bureau of Labor Statistics (BLS) API client for employment and wage data."""

import logging
from datetime import datetime
import requests

logger = logging.getLogger(__name__)

BASE_URL = "https://api.bls.gov/publicAPI/v2/timeseries/data/"
TIMEOUT = 15


class BLSClient:
    """Client for the Bureau of Labor Statistics API."""

    def __init__(self):
        self.session = requests.Session()

    def _fetch_series(self, series_ids: list[str],
                      start_year: int = None, end_year: int = None) -> dict:
        """Fetch time series data from BLS."""
        try:
            current_year = datetime.now().year
            if not end_year:
                end_year = current_year
            if not start_year:
                start_year = current_year - 2

            payload = {
                "seriesid": series_ids,
                "startyear": str(start_year),
                "endyear": str(end_year),
            }

            resp = self.session.post(BASE_URL, json=payload, timeout=TIMEOUT)
            resp.raise_for_status()
            data = resp.json()

            if data.get("status") != "REQUEST_SUCCEEDED":
                return {"error": data.get("message", "BLS request failed"), "series": {}}

            result = {}
            for series in data.get("Results", {}).get("series", []):
                sid = series["seriesID"]
                observations = []
                for item in series.get("data", []):
                    try:
                        observations.append({
                            "year": int(item["year"]),
                            "period": item["period"],
                            "value": float(item["value"]),
                            "date": f"{item['year']}-{item['period'][1:]}",
                        })
                    except (ValueError, KeyError):
                        continue
                result[sid] = observations

            return {"series": result}

        except Exception as e:
            logger.warning(f"BLS API error: {e}")
            return {"error": str(e), "series": {}}

    def get_employment_data(self, state_fips: str = "00") -> dict:
        """Get total nonfarm employment.
        state_fips='00' means national. State codes: TX='48', CA='06', etc.
        """
        # CES series: Total nonfarm employment (national)
        series_id = "CES0000000001"
        result = self._fetch_series([series_id])
        data = result.get("series", {}).get(series_id, [])

        latest = data[0] if data else None
        return {
            "total_employment_thousands": latest["value"] if latest else None,
            "latest_period": latest["date"] if latest else None,
            "recent_data": data[:12],
            "source": "BLS (CES0000000001)",
            "error": result.get("error"),
        }

    def get_unemployment_rate(self) -> dict:
        """Get national unemployment rate from BLS."""
        # LNS14000000 = Unemployment Rate
        series_id = "LNS14000000"
        result = self._fetch_series([series_id])
        data = result.get("series", {}).get(series_id, [])

        latest = data[0] if data else None
        return {
            "unemployment_rate": latest["value"] if latest else None,
            "latest_period": latest["date"] if latest else None,
            "recent_data": data[:12],
            "source": "BLS (LNS14000000)",
            "error": result.get("error"),
        }

    def get_cpi_urban(self) -> dict:
        """Get CPI-U (All Urban Consumers)."""
        # CUSR0000SA0 = CPI-U All items
        series_id = "CUSR0000SA0"
        result = self._fetch_series([series_id])
        data = result.get("series", {}).get(series_id, [])

        latest = data[0] if data else None
        yoy_change = None
        if len(data) >= 13:
            current = data[0]["value"]
            year_ago = data[12]["value"]
            yoy_change = round(((current - year_ago) / year_ago) * 100, 2)

        return {
            "latest_cpi": latest["value"] if latest else None,
            "yoy_inflation": yoy_change,
            "recent_data": data[:12],
            "source": "BLS (CUSR0000SA0)",
            "error": result.get("error"),
        }

    def get_all_labor_data(self) -> dict:
        """Fetch all labor market data in one call."""
        series_ids = [
            "CES0000000001",  # Total nonfarm employment
            "LNS14000000",    # Unemployment rate
            "CUSR0000SA0",    # CPI-U
        ]
        result = self._fetch_series(series_ids)
        series = result.get("series", {})

        emp_data = series.get("CES0000000001", [])
        unemp_data = series.get("LNS14000000", [])
        cpi_data = series.get("CUSR0000SA0", [])

        return {
            "employment": {
                "total_thousands": emp_data[0]["value"] if emp_data else None,
                "recent_data": emp_data[:6],
                "source": "BLS (CES0000000001)",
            },
            "unemployment": {
                "rate": unemp_data[0]["value"] if unemp_data else None,
                "recent_data": unemp_data[:6],
                "source": "BLS (LNS14000000)",
            },
            "cpi": {
                "latest": cpi_data[0]["value"] if cpi_data else None,
                "recent_data": cpi_data[:6],
                "source": "BLS (CUSR0000SA0)",
            },
            "error": result.get("error"),
        }
