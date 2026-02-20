"""Walk Score API client for walkability, transit, and bike scores.

Walk Score API: https://www.walkscore.com/professional/api.php
Free tier: 5,000 requests/day.

Scores:
  Walk Score  0-49  Car-Dependent / 50-69 Somewhat Walkable /
              70-89 Very Walkable / 90-100 Walker's Paradise
  Transit     0-24  Minimal / 25-49 Some / 50-69 Excellent /
              70-89 Excellent / 90-100 Rider's Paradise
  Bike        0-49  Bikeable / 50-69 Bikeable / 70-84 Very Bikeable /
              85-100 Biker's Paradise
"""

import logging
import requests

from config import WALKSCORE_API_KEY

logger = logging.getLogger(__name__)

BASE_URL = "https://api.walkscore.com/score"
TIMEOUT = 10

WALK_LABELS = [
    (90, "Walker's Paradise"),
    (70, "Very Walkable"),
    (50, "Somewhat Walkable"),
    (25, "Car-Dependent"),
    (0,  "Car-Dependent"),
]

TRANSIT_LABELS = [
    (90, "Rider's Paradise"),
    (70, "Excellent Transit"),
    (50, "Excellent Transit"),
    (25, "Some Transit"),
    (0,  "Minimal Transit"),
]

BIKE_LABELS = [
    (85, "Biker's Paradise"),
    (70, "Very Bikeable"),
    (50, "Bikeable"),
    (0,  "Bikeable"),
]


def _label(score: int, thresholds: list) -> str:
    if score is None:
        return "N/A"
    for threshold, label in thresholds:
        if score >= threshold:
            return label
    return "N/A"


class WalkScoreClient:
    """Client for the Walk Score API."""

    def __init__(self, api_key: str = None):
        self.api_key = api_key or WALKSCORE_API_KEY

    def _is_configured(self) -> bool:
        return bool(self.api_key)

    def get_scores(self, address: str, lat: float = None,
                   lon: float = None) -> dict:
        """Fetch Walk Score, Transit Score, and Bike Score for an address.

        Args:
            address: Full property address string
            lat: Latitude (improves accuracy and avoids geocoding delay)
            lon: Longitude

        Returns:
            Dict with walk_score, transit_score, bike_score, labels, and
            an underwriting risk note.
        """
        if not self._is_configured():
            return {
                "available": False,
                "note": "WALKSCORE_API_KEY not configured. "
                        "Register at walkscore.com/professional/api.php",
            }

        params = {
            "format": "json",
            "wsapikey": self.api_key,
            "address": address,
            "transit": 1,
            "bike": 1,
        }
        if lat is not None and lon is not None:
            params["lat"] = lat
            params["lon"] = lon

        try:
            resp = requests.get(BASE_URL, params=params, timeout=TIMEOUT)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            logger.warning(f"Walk Score API error: {e}")
            return {"available": False, "error": str(e)}

        status = data.get("status")
        if status not in (1, 2):  # 1=score, 2=score+description
            return {
                "available": False,
                "status_code": status,
                "note": _status_note(status),
            }

        walk = data.get("walkscore")
        transit_data = data.get("transit", {}) or {}
        bike_data = data.get("bike", {}) or {}

        transit = transit_data.get("score")
        bike = bike_data.get("score")

        return {
            "available": True,
            "address": data.get("snapped_address", address),
            "walk_score": walk,
            "walk_label": _label(walk, WALK_LABELS),
            "transit_score": transit,
            "transit_label": _label(transit, TRANSIT_LABELS),
            "bike_score": bike,
            "bike_label": _label(bike, BIKE_LABELS),
            "walkscore_url": data.get("ws_link", ""),
            "underwriting_note": _underwriting_note(walk, transit),
            "source": "Walk Score API",
        }


def _status_note(status: int) -> str:
    notes = {
        0: "Internal error from Walk Score",
        30: "Invalid API key",
        31: "Walk Score daily API quota exceeded",
        40: "Service area not supported for this address",
        41: "Coordinates not provided for a score-only request",
    }
    return notes.get(status, f"Walk Score status {status}")


def _underwriting_note(walk: int, transit: int) -> str:
    if walk is None:
        return ""
    parts = []
    if walk >= 70:
        parts.append("High walkability supports premium rent positioning and broad tenant appeal.")
    elif walk >= 50:
        parts.append("Moderate walkability — amenity access may influence tenant quality.")
    else:
        parts.append("Low walkability increases car-dependency; factor into tenant demand forecasts.")

    if transit is not None:
        if transit >= 70:
            parts.append("Excellent transit access expands renter pool and supports rent growth.")
        elif transit < 25:
            parts.append("Limited transit may constrain demand from non-car-owning renters.")
    return " ".join(parts)
