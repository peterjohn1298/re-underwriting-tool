"""RentCast API client for real rental market data, rent estimates, and comps.

RentCast provides:
  - /avm/rent/long-term  : AVM rent estimate for a specific address
  - /markets             : Market-level stats by zip code (avg rent, days on market, vacancy)
  - /listings/rental/long-term : Real rental comps near an address
"""

import logging
import requests

from config import RENTCAST_API_KEY

logger = logging.getLogger(__name__)

BASE_URL = "https://api.rentcast.io/v1"
TIMEOUT = 15

# Map our internal property types to RentCast property types
PROPERTY_TYPE_MAP = {
    "Multifamily": "Apartment",
    "Single Family": "Single Family",
    "Condo": "Condo",
    "Townhouse": "Townhouse",
    # Commercial types not well supported by RentCast — fall back gracefully
    "Office": None,
    "Retail": None,
    "Industrial": None,
}


class RentCastClient:
    """Client for the RentCast rental market data API."""

    def __init__(self, api_key: str = None):
        self.api_key = api_key or RENTCAST_API_KEY
        self.session = requests.Session()
        if self.api_key:
            self.session.headers.update({
                "X-Api-Key": self.api_key,
                "Accept": "application/json",
            })

    def _is_configured(self) -> bool:
        return bool(self.api_key)

    def _get(self, endpoint: str, params: dict) -> dict:
        """Make a GET request to the RentCast API."""
        if not self._is_configured():
            return {"error": "RENTCAST_API_KEY not configured"}
        try:
            url = f"{BASE_URL}/{endpoint.lstrip('/')}"
            resp = self.session.get(url, params=params, timeout=TIMEOUT)
            if resp.status_code == 429:
                return {"error": "RentCast rate limit reached — upgrade plan or wait"}
            if resp.status_code == 404:
                return {"error": "No data found for this location"}
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            logger.warning(f"RentCast API error ({endpoint}): {e}")
            return {"error": str(e)}

    def get_market_stats(self, zip_code: str, property_type: str = "Apartment",
                         bedrooms: int = None, history_range: int = 12) -> dict:
        """Get rental market statistics for a zip code.

        Args:
            zip_code: 5-digit zip code
            property_type: RentCast property type (Apartment, Single Family, etc.)
            bedrooms: Optional bedroom filter
            history_range: Months of history to return (max 12 on free tier)

        Returns:
            Dict with averageRent, medianRent, minRent, maxRent,
            averageDaysOnMarket, totalListings, and history.
        """
        params = {
            "zipCode": zip_code,
            "propertyType": property_type,
            "historyRange": history_range,
        }
        if bedrooms is not None:
            params["bedrooms"] = bedrooms

        data = self._get("/markets", params)
        if "error" in data:
            return {"error": data["error"], "available": False}

        # Data is nested under rentalData key
        rd = data.get("rentalData", {})

        # Try to find property-type-specific stats from dataByPropertyType
        pt_stats = {}
        for pt_entry in rd.get("dataByPropertyType", []):
            if pt_entry.get("propertyType", "").lower() == property_type.lower():
                pt_stats = pt_entry
                break

        # Use property-type-specific stats if found, else use overall rentalData
        def _get_stat(key):
            return pt_stats.get(key) or rd.get(key)

        # Compute annualized rent trend from history
        rent_trend = None
        history = rd.get("history", {})
        if len(history) >= 2:
            sorted_months = sorted(history.keys())
            first_month = history[sorted_months[0]]
            last_month = history[sorted_months[-1]]
            # Use property-type history if available
            def _hist_rent(month_data):
                for pt in month_data.get("dataByPropertyType", []):
                    if pt.get("propertyType", "").lower() == property_type.lower():
                        return pt.get("averageRent")
                return month_data.get("averageRent")
            first_rent = _hist_rent(first_month)
            last_rent = _hist_rent(last_month)
            if first_rent and last_rent and first_rent > 0:
                n_months = len(sorted_months) - 1
                if n_months > 0:
                    total_change = (last_rent - first_rent) / first_rent * 100
                    annualized = total_change * (12 / n_months)
                    rent_trend = round(annualized, 2)

        return {
            "available": True,
            "zip_code": zip_code,
            "property_type": property_type,
            "average_rent": _get_stat("averageRent"),
            "median_rent": _get_stat("medianRent"),
            "min_rent": _get_stat("minRent"),
            "max_rent": _get_stat("maxRent"),
            "avg_days_on_market": _get_stat("averageDaysOnMarket"),
            "total_listings": _get_stat("totalListings"),
            "avg_sq_ft": _get_stat("averageSquareFootage"),
            "rent_per_sqft": _get_stat("averageRentPerSquareFoot"),
            "annualized_rent_trend_pct": rent_trend,
            "data_by_bedrooms": rd.get("dataByBedrooms", []),
            "source": "RentCast",
        }

    def get_rent_estimate(self, address: str, property_type: str = "Apartment",
                          bedrooms: int = 2, bathrooms: float = 1.0,
                          sq_ft: int = None, comp_count: int = 5) -> dict:
        """Get AVM rent estimate for a specific address.

        Args:
            address: Full property address (e.g. "123 Main St, Austin, TX 78701")
            property_type: RentCast property type
            bedrooms: Number of bedrooms
            bathrooms: Number of bathrooms
            sq_ft: Square footage (improves accuracy)
            comp_count: Number of comparable listings to return

        Returns:
            Dict with estimated rent, range, and comparable listings used.
        """
        params = {
            "address": address,
            "propertyType": property_type,
            "bedrooms": bedrooms,
            "bathrooms": bathrooms,
            "compCount": comp_count,
        }
        if sq_ft:
            params["squareFootage"] = sq_ft

        data = self._get("/avm/rent/long-term", params)
        if "error" in data:
            return {"error": data["error"], "available": False}

        comps = []
        for listing in data.get("listings", []):
            comps.append({
                "address": listing.get("formattedAddress", ""),
                "city": listing.get("city", ""),
                "state": listing.get("state", ""),
                "zip": listing.get("zipCode", ""),
                "bedrooms": listing.get("bedrooms"),
                "bathrooms": listing.get("bathrooms"),
                "sq_ft": listing.get("squareFootage"),
                "rent": listing.get("price"),
                "days_on_market": listing.get("daysOnMarket"),
                "distance_miles": round(listing.get("distance", 0), 2),
                "source": "RentCast",
                "synthetic": False,
            })

        return {
            "available": True,
            "address": address,
            "estimated_rent": data.get("rent"),
            "rent_range_low": data.get("rentRangeLow"),
            "rent_range_high": data.get("rentRangeHigh"),
            "latitude": data.get("latitude"),
            "longitude": data.get("longitude"),
            "comps_used": comps,
            "source": "RentCast AVM",
        }

    def get_rental_comps(self, address: str, radius_miles: float = 1.0,
                         property_type: str = "Apartment", bedrooms: int = None,
                         limit: int = 10) -> dict:
        """Get real active rental listings near an address.

        Args:
            address: Full property address
            radius_miles: Search radius in miles
            property_type: RentCast property type
            bedrooms: Optional bedroom filter
            limit: Max number of listings to return

        Returns:
            Dict with list of real comparable rental listings.
        """
        params = {
            "address": address,
            "radius": radius_miles,
            "propertyType": property_type,
            "status": "Active",
            "limit": limit,
        }
        if bedrooms is not None:
            params["bedrooms"] = bedrooms

        data = self._get("/listings/rental/long-term", params)
        if isinstance(data, dict) and "error" in data:
            return {"error": data["error"], "available": False, "listings": []}

        # RentCast returns a list directly for this endpoint
        listings_raw = data if isinstance(data, list) else data.get("data", [])
        listings = []
        for item in listings_raw:
            listings.append({
                "address": item.get("formattedAddress", ""),
                "city": item.get("city", ""),
                "state": item.get("state", ""),
                "zip": item.get("zipCode", ""),
                "bedrooms": item.get("bedrooms"),
                "bathrooms": item.get("bathrooms"),
                "sq_ft": item.get("squareFootage"),
                "rent": item.get("price"),
                "days_on_market": item.get("daysOnMarket"),
                "listed_date": item.get("listedDate", ""),
                "source": "RentCast",
                "synthetic": False,
            })

        avg_rent = None
        rents = [l["rent"] for l in listings if l["rent"]]
        if rents:
            avg_rent = round(sum(rents) / len(rents))

        return {
            "available": True,
            "listings": listings,
            "count": len(listings),
            "average_rent": avg_rent,
            "radius_miles": radius_miles,
            "source": "RentCast",
        }

    def get_all_rental_data(self, address: str, zip_code: str,
                            property_type: str, avg_bedrooms: int = 2,
                            avg_sq_ft: int = None) -> dict:
        """Convenience method — fetches market stats, AVM, and comps in one call.

        Args:
            address: Full property address
            zip_code: Property zip code
            property_type: Our internal property type (e.g. 'Multifamily')
            avg_bedrooms: Average unit bedroom count
            avg_sq_ft: Average unit square footage

        Returns:
            Combined dict with market_stats, rent_estimate, and rental_comps.
        """
        rc_type = PROPERTY_TYPE_MAP.get(property_type, "Apartment")
        if rc_type is None:
            return {
                "available": False,
                "note": f"RentCast does not support {property_type} property type",
            }

        market_stats = self.get_market_stats(zip_code, rc_type) if zip_code else {"available": False}
        rent_estimate = self.get_rent_estimate(address, rc_type, avg_bedrooms,
                                               sq_ft=avg_sq_ft) if address else {"available": False}
        rental_comps = self.get_rental_comps(address, property_type=rc_type,
                                             bedrooms=avg_bedrooms) if address else {"available": False, "listings": []}

        return {
            "available": True,
            "market_stats": market_stats,
            "rent_estimate": rent_estimate,
            "rental_comps": rental_comps,
            "property_type": rc_type,
            "source": "RentCast",
        }
