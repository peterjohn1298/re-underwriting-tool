"""OpenFEMA NFHL Flood Zone client.

Uses the FEMA National Flood Hazard Layer (NFHL) ArcGIS REST API to look up
the flood zone designation for a given latitude/longitude.

No API key required. Free and public.
Endpoint: https://hazards.fema.gov/gis/nfhl/rest/services/public/NFHL/MapServer/28/query
"""

import logging
import requests

logger = logging.getLogger(__name__)

_NFHL_URL = (
    "https://hazards.fema.gov/gis/nfhl/rest/services/public/NFHL/MapServer/28/query"
)

# FEMA flood zone risk descriptions
_ZONE_DESCRIPTIONS = {
    "A":   ("High Risk", "Special Flood Hazard Area — 1% annual chance of flooding. Flood insurance typically required by lenders."),
    "AE":  ("High Risk", "Special Flood Hazard Area with base flood elevations. 1% annual chance. Insurance required."),
    "AH":  ("High Risk", "1% annual chance of shallow flooding (ponding), depths 1–3 ft."),
    "AO":  ("High Risk", "1% annual chance river/stream flood — sheet flow, 1–3 ft depth."),
    "AR":  ("High Risk", "Area of temporary increased flood risk — levee under construction."),
    "A99": ("High Risk", "Area protected by a Federal flood control system under construction."),
    "VE":  ("Very High Risk", "Coastal high-hazard area — wave action likely. Highest risk category."),
    "V":   ("Very High Risk", "Coastal high-hazard area without base flood elevations."),
    "X":   ("Minimal Risk", "Area of minimal flood hazard. Outside 0.2% annual chance floodplain."),
    "B":   ("Moderate Risk", "Moderate flood hazard — between 0.2% and 1% annual chance. (Pre-FIRM zones)"),
    "C":   ("Minimal Risk",  "Minimal flood hazard area. (Pre-FIRM zones)"),
    "D":   ("Undetermined", "Area of undetermined but possible flood hazard."),
}


def _zone_risk(zone: str) -> tuple[str, str]:
    """Return (risk_level, description) for a flood zone code."""
    if not zone:
        return "Unknown", "Flood zone not determined."
    # Try exact match, then prefix match
    clean = zone.upper().strip()
    if clean in _ZONE_DESCRIPTIONS:
        return _ZONE_DESCRIPTIONS[clean]
    for key, val in _ZONE_DESCRIPTIONS.items():
        if clean.startswith(key):
            return val
    return "Unknown", f"Flood zone {zone} — check FEMA Flood Map Service Center."


def get_flood_zone(lat: float, lon: float) -> dict:
    """Look up the FEMA flood zone for a lat/lon coordinate.

    Returns a dict with keys:
        available (bool), flood_zone (str), risk_level (str),
        description (str), sfha (bool), source (str)
    """
    if not lat or not lon:
        return {"available": False, "note": "No coordinates provided"}

    try:
        params = {
            "geometry": f"{lon},{lat}",
            "geometryType": "esriGeometryPoint",
            "spatialRel": "esriSpatialRelIntersects",
            "outFields": "FLD_ZONE,ZONE_SUBTY,SFHA_TF,STUDY_TYP",
            "returnGeometry": "false",
            "f": "json",
            "inSR": "4326",
            "outSR": "4326",
        }
        resp = requests.get(_NFHL_URL, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()

        features = data.get("features", [])
        if not features:
            return {
                "available": True,
                "flood_zone": "X",
                "risk_level": "Minimal Risk",
                "description": "No Special Flood Hazard Area found at this location. Likely Zone X (minimal risk).",
                "sfha": False,
                "source": "FEMA NFHL REST API",
                "note": "No NFHL polygon found — property may be outside mapped floodplain.",
            }

        attrs = features[0].get("attributes", {})
        zone = attrs.get("FLD_ZONE", "X")
        sfha_raw = attrs.get("SFHA_TF", "F")
        sfha = str(sfha_raw).upper() in ("T", "TRUE", "1", "YES")
        subtype = attrs.get("ZONE_SUBTY", "")
        risk_level, description = _zone_risk(zone)

        return {
            "available": True,
            "flood_zone": zone,
            "zone_subtype": subtype or None,
            "risk_level": risk_level,
            "description": description,
            "sfha": sfha,
            "source": "FEMA NFHL REST API (OpenFEMA)",
            "lat": lat,
            "lon": lon,
        }

    except requests.exceptions.Timeout:
        logger.warning("FEMA NFHL API timeout")
        return {"available": False, "note": "FEMA API timeout"}
    except Exception as e:
        logger.warning(f"FEMA flood zone lookup failed: {e}")
        return {"available": False, "note": str(e)}
