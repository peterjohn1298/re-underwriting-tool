"""Zillow ZORI (Observed Rent Index) client using local CSV data.

Zillow publishes ZORI as a free public CSV. This client reads the local copy
and provides city-level rent trends and annual growth rates.
"""

import os
import logging
import pandas as pd

from config import DATA_DIR

logger = logging.getLogger(__name__)

ZORI_CSV = os.path.join(DATA_DIR, "zori_sm_sa_month.csv")

# Map common city names to ZORI RegionName
CITY_ALIASES = {
    "New York City": "New York",
    "NYC": "New York",
    "LA": "Los Angeles",
    "SF": "San Francisco",
    "DC": "Washington",
    "D.C.": "Washington",
    "Washington DC": "Washington",
    "Washington D.C.": "Washington",
    "St Louis": "St. Louis",
    "Saint Louis": "St. Louis",
}

# Map state abbreviation to full names used in CSV
STATE_MAP = {
    "TX": "TX", "CA": "CA", "NY": "NY", "FL": "FL", "IL": "IL",
    "PA": "PA", "GA": "GA", "NC": "NC", "OH": "OH", "MN": "MN",
    "CO": "CO", "AZ": "AZ", "NV": "NV", "OR": "OR", "WA": "WA",
    "MA": "MA", "MD": "MD", "MO": "MO", "IN": "IN", "TN": "TN",
    "UT": "UT", "DC": "DC",
}


class ZillowClient:
    """Client for Zillow ZORI rent data from local CSV."""

    def __init__(self):
        self._df = None
        self._load()

    def _load(self):
        """Load the ZORI CSV into a DataFrame."""
        if not os.path.exists(ZORI_CSV):
            logger.warning(f"ZORI CSV not found at {ZORI_CSV}")
            return

        try:
            self._df = pd.read_csv(ZORI_CSV)
            logger.info(f"ZORI data loaded: {len(self._df)} metro areas")
        except Exception as e:
            logger.error(f"Failed to load ZORI CSV: {e}")
            self._df = None

    def _find_city(self, city: str, state: str = "") -> pd.Series | None:
        """Find a city row in the ZORI data."""
        if self._df is None:
            return None

        # Normalize
        city_clean = city.strip().title()
        city_clean = CITY_ALIASES.get(city_clean, city_clean)

        # Try exact match on RegionName
        matches = self._df[self._df["RegionName"].str.lower() == city_clean.lower()]
        if len(matches) == 1:
            return matches.iloc[0]
        if len(matches) > 1 and state:
            state_matches = matches[matches["StateName"].str.upper() == state.upper()]
            if len(state_matches) >= 1:
                return state_matches.iloc[0]
            return matches.iloc[0]

        # Try contains
        contains = self._df[self._df["RegionName"].str.lower().str.contains(city_clean.lower())]
        if len(contains) >= 1:
            return contains.iloc[0]

        return None

    def get_city_rent_trend(self, city: str, state: str = "") -> dict:
        """Get rent trend data for a city.

        Returns dict with time series of monthly rents and metadata.
        """
        row = self._find_city(city, state)
        if row is None:
            return {"error": f"No ZORI data found for {city}, {state}", "available": False}

        # Extract date columns (format: YYYY-MM)
        date_cols = [c for c in self._df.columns if len(c) == 7 and c[4] == "-"]
        rents = {}
        for col in date_cols:
            val = row.get(col)
            if pd.notna(val):
                rents[col] = float(val)

        if not rents:
            return {"error": "No rent data available", "available": False}

        dates = sorted(rents.keys())
        return {
            "available": True,
            "city": row["RegionName"],
            "state": row.get("StateName", state),
            "region_type": row.get("RegionType", "msa"),
            "latest_rent": rents[dates[-1]],
            "earliest_rent": rents[dates[0]],
            "rent_series": {d: rents[d] for d in dates},
            "data_points": len(rents),
            "source": "Zillow ZORI (Smoothed, Seasonally Adjusted)",
        }

    def get_annual_growth_rates(self, city: str, state: str = "") -> dict:
        """Calculate annual rent growth rates from ZORI data.

        Returns dict with annual growth rates suitable for blending
        into the rent predictor.
        """
        trend = self.get_city_rent_trend(city, state)
        if not trend.get("available"):
            return {"error": trend.get("error", "No data"), "available": False, "growth_rates": []}

        series = trend["rent_series"]
        dates = sorted(series.keys())

        # Calculate year-over-year growth for each January
        jan_dates = [d for d in dates if d.endswith("-01")]
        annual_rates = []
        for i in range(1, len(jan_dates)):
            prev = series[jan_dates[i - 1]]
            curr = series[jan_dates[i]]
            if prev > 0:
                rate = ((curr - prev) / prev) * 100
                annual_rates.append({
                    "year": jan_dates[i][:4],
                    "growth_rate": round(rate, 2),
                    "rent": round(curr, 0),
                })

        # Also calculate growth between available 6-month intervals
        six_month_rates = []
        for i in range(1, len(dates)):
            prev = series[dates[i - 1]]
            curr = series[dates[i]]
            if prev > 0:
                # Annualize based on month gap
                d_prev = dates[i - 1]
                d_curr = dates[i]
                prev_months = int(d_prev[:4]) * 12 + int(d_prev[5:7])
                curr_months = int(d_curr[:4]) * 12 + int(d_curr[5:7])
                gap = curr_months - prev_months
                if gap > 0:
                    period_rate = (curr - prev) / prev
                    annualized = ((1 + period_rate) ** (12 / gap) - 1) * 100
                    six_month_rates.append(round(annualized, 2))

        avg_growth = round(sum(r["growth_rate"] for r in annual_rates) / len(annual_rates), 2) if annual_rates else None

        return {
            "available": True,
            "city": trend["city"],
            "state": trend["state"],
            "annual_rates": annual_rates,
            "avg_annual_growth": avg_growth,
            "latest_rent": trend["latest_rent"],
            "annualized_rates": six_month_rates[-6:] if six_month_rates else [],
            "source": "Zillow ZORI (Smoothed, Seasonally Adjusted)",
        }
