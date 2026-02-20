"""Unit mix model — Studio / 1BR / 2BR / 3BR breakdown.

Computes blended rent and per-type revenue for use in the financial model.
When a unit mix is provided, the blended rent overrides the single
in_place_rent input so the pro forma reflects the actual unit composition.
"""

from dataclasses import dataclass, field
from typing import List


UNIT_TYPES = ["Studio", "1BR", "2BR", "3BR"]

# HUD bedroom keys for FMR comparison
HUD_BEDROOM_KEYS = {
    "Studio": "efficiency",
    "1BR": "one_br",
    "2BR": "two_br",
    "3BR": "three_br",
}


@dataclass
class UnitType:
    """A single bedroom-type slice of the unit mix."""
    type_name: str          # "Studio", "1BR", "2BR", "3BR"
    count: int              # number of units of this type
    in_place_rent: float    # $/unit/month currently in place
    market_rent: float = 0.0  # $/unit/month market rate (optional)

    @property
    def annual_revenue(self) -> float:
        return self.count * self.in_place_rent * 12

    @property
    def annual_market_revenue(self) -> float:
        return self.count * (self.market_rent or self.in_place_rent) * 12

    @property
    def rent_premium_per_unit(self) -> float:
        return (self.market_rent - self.in_place_rent) if self.market_rent > 0 else 0.0

    def to_dict(self) -> dict:
        return {
            "type_name": self.type_name,
            "count": self.count,
            "in_place_rent": self.in_place_rent,
            "market_rent": self.market_rent,
            "annual_revenue": self.annual_revenue,
            "annual_market_revenue": self.annual_market_revenue,
            "rent_premium_per_unit": self.rent_premium_per_unit,
        }


@dataclass
class UnitMix:
    """Full unit mix for a property."""
    units: List[UnitType] = field(default_factory=list)

    def is_empty(self) -> bool:
        return not self.units or all(u.count == 0 for u in self.units)

    @property
    def total_units(self) -> int:
        return sum(u.count for u in self.units)

    @property
    def blended_in_place_rent(self) -> float:
        """Weighted average in-place rent per unit per month."""
        total = self.total_units
        if total == 0:
            return 0.0
        return sum(u.in_place_rent * u.count for u in self.units) / total

    @property
    def blended_market_rent(self) -> float:
        """Weighted average market rent per unit per month."""
        total = self.total_units
        if total == 0:
            return 0.0
        return sum((u.market_rent or u.in_place_rent) * u.count for u in self.units) / total

    @property
    def total_annual_revenue(self) -> float:
        return sum(u.annual_revenue for u in self.units)

    @property
    def total_annual_market_revenue(self) -> float:
        return sum(u.annual_market_revenue for u in self.units)

    @property
    def total_rent_premium_potential(self) -> float:
        """Total annual revenue upside if all units reach market rent."""
        return self.total_annual_market_revenue - self.total_annual_revenue

    def to_dict(self) -> dict:
        return {
            "units": [u.to_dict() for u in self.units],
            "total_units": self.total_units,
            "blended_in_place_rent": round(self.blended_in_place_rent, 2),
            "blended_market_rent": round(self.blended_market_rent, 2),
            "total_annual_revenue": round(self.total_annual_revenue, 2),
            "total_annual_market_revenue": round(self.total_annual_market_revenue, 2),
            "total_rent_premium_potential": round(self.total_rent_premium_potential, 2),
        }

    def with_hud_fmr(self, hud_fmr: dict) -> dict:
        """Annotate each unit type with its HUD FMR comparison.

        Args:
            hud_fmr: the hud_fmr dict from market_data (must have fmr_by_bedroom key)

        Returns:
            dict with units enriched with hud_fmr and vs_hud_pct fields
        """
        fmr_map = hud_fmr.get("fmr_by_bedroom", {}) if hud_fmr else {}
        result = self.to_dict()
        for u_dict in result["units"]:
            hud_key = HUD_BEDROOM_KEYS.get(u_dict["type_name"])
            fmr = fmr_map.get(hud_key)
            if fmr and isinstance(fmr, (int, float)) and fmr > 0:
                u_dict["hud_fmr"] = fmr
                u_dict["vs_hud_pct"] = round(
                    (u_dict["in_place_rent"] - fmr) / fmr * 100, 1
                )
            else:
                u_dict["hud_fmr"] = None
                u_dict["vs_hud_pct"] = None
        return result


def parse_unit_mix_from_form(form) -> UnitMix:
    """Parse unit mix inputs from Flask form data.

    Expects form fields:
        um_count_Studio, um_rent_Studio, um_market_rent_Studio
        um_count_1BR,    um_rent_1BR,    um_market_rent_1BR
        um_count_2BR,    um_rent_2BR,    um_market_rent_2BR
        um_count_3BR,    um_rent_3BR,    um_market_rent_3BR
    """
    def _i(v):
        try:
            return int(v) if v else 0
        except (ValueError, TypeError):
            return 0

    def _f(v):
        try:
            return float(v) if v else 0.0
        except (ValueError, TypeError):
            return 0.0

    units = []
    for t in UNIT_TYPES:
        count = _i(form.get(f"um_count_{t}"))
        rent = _f(form.get(f"um_rent_{t}"))
        market = _f(form.get(f"um_market_rent_{t}"))
        if count > 0 and rent > 0:
            units.append(UnitType(
                type_name=t,
                count=count,
                in_place_rent=rent,
                market_rent=market,
            ))

    return UnitMix(units=units)
