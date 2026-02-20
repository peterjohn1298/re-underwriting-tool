"""Tests for unit mix model."""
import pytest
from models.unit_mix import UnitType, UnitMix, parse_unit_mix_from_form, HUD_BEDROOM_KEYS


# ---------------------------------------------------------------------------
# UnitType tests
# ---------------------------------------------------------------------------

class TestUnitType:
    def test_annual_revenue(self):
        u = UnitType("1BR", count=10, in_place_rent=1_500)
        assert u.annual_revenue == pytest.approx(10 * 1_500 * 12)

    def test_rent_premium_per_unit_positive(self):
        u = UnitType("2BR", count=5, in_place_rent=1_800, market_rent=2_000)
        assert u.rent_premium_per_unit == pytest.approx(200.0)

    def test_rent_premium_zero_when_no_market_rent(self):
        u = UnitType("2BR", count=5, in_place_rent=1_800, market_rent=0.0)
        assert u.rent_premium_per_unit == 0.0

    def test_annual_market_revenue_uses_market_rent(self):
        u = UnitType("2BR", count=10, in_place_rent=1_800, market_rent=2_000)
        assert u.annual_market_revenue == pytest.approx(10 * 2_000 * 12)

    def test_annual_market_revenue_falls_back_to_in_place(self):
        u = UnitType("2BR", count=10, in_place_rent=1_800, market_rent=0.0)
        assert u.annual_market_revenue == pytest.approx(10 * 1_800 * 12)

    def test_to_dict_has_required_keys(self):
        u = UnitType("Studio", count=5, in_place_rent=1_200)
        d = u.to_dict()
        for key in ("type_name", "count", "in_place_rent", "market_rent",
                    "annual_revenue", "annual_market_revenue", "rent_premium_per_unit"):
            assert key in d


# ---------------------------------------------------------------------------
# UnitMix tests
# ---------------------------------------------------------------------------

class TestUnitMix:
    def _standard_mix(self):
        """100-unit mix: 10 Studio / 30 1BR / 50 2BR / 10 3BR."""
        return UnitMix(units=[
            UnitType("Studio", 10,  1_100, 1_200),
            UnitType("1BR",    30,  1_500, 1_600),
            UnitType("2BR",    50,  1_850, 2_000),
            UnitType("3BR",    10,  2_200, 2_400),
        ])

    def test_total_units(self):
        mix = self._standard_mix()
        assert mix.total_units == 100

    def test_blended_rent_known_answer(self):
        mix = self._standard_mix()
        expected = (10*1_100 + 30*1_500 + 50*1_850 + 10*2_200) / 100
        assert mix.blended_in_place_rent == pytest.approx(expected, rel=1e-3)

    def test_blended_market_rent_higher_than_in_place(self):
        mix = self._standard_mix()
        assert mix.blended_market_rent > mix.blended_in_place_rent

    def test_total_annual_revenue(self):
        mix = self._standard_mix()
        expected = (10*1_100 + 30*1_500 + 50*1_850 + 10*2_200) * 12
        assert mix.total_annual_revenue == pytest.approx(expected, rel=1e-3)

    def test_rent_premium_potential_positive(self):
        mix = self._standard_mix()
        assert mix.total_rent_premium_potential > 0

    def test_is_empty_false_for_standard_mix(self):
        mix = self._standard_mix()
        assert mix.is_empty() is False

    def test_is_empty_true_for_zero_counts(self):
        mix = UnitMix(units=[UnitType("1BR", 0, 1_500)])
        assert mix.is_empty() is True

    def test_is_empty_true_for_no_units(self):
        mix = UnitMix(units=[])
        assert mix.is_empty() is True

    def test_blended_rent_zero_for_empty_mix(self):
        mix = UnitMix(units=[])
        assert mix.blended_in_place_rent == 0.0

    def test_to_dict_has_required_keys(self):
        mix = self._standard_mix()
        d = mix.to_dict()
        for key in ("units", "total_units", "blended_in_place_rent",
                    "blended_market_rent", "total_annual_revenue",
                    "total_annual_market_revenue", "total_rent_premium_potential"):
            assert key in d


# ---------------------------------------------------------------------------
# HUD FMR annotation
# ---------------------------------------------------------------------------

class TestWithHudFmr:
    def _mix(self):
        return UnitMix(units=[
            UnitType("Studio", 5, 1_100),
            UnitType("1BR",   10, 1_500),
            UnitType("2BR",   20, 1_850),
            UnitType("3BR",    5, 2_200),
        ])

    def _hud(self):
        return {
            "fmr_by_bedroom": {
                "efficiency": 1_000,
                "one_br":     1_400,
                "two_br":     1_800,
                "three_br":   2_100,
            }
        }

    def test_annotates_all_unit_types(self):
        result = self._mix().with_hud_fmr(self._hud())
        for u in result["units"]:
            assert u["hud_fmr"] is not None

    def test_vs_hud_pct_correct_for_studio(self):
        result = self._mix().with_hud_fmr(self._hud())
        studio = result["units"][0]
        expected = round((1_100 - 1_000) / 1_000 * 100, 1)
        assert studio["vs_hud_pct"] == pytest.approx(expected)

    def test_no_crash_when_hud_unavailable(self):
        result = self._mix().with_hud_fmr(None)
        for u in result["units"]:
            assert u["hud_fmr"] is None
            assert u["vs_hud_pct"] is None

    def test_no_crash_when_hud_missing_bedroom(self):
        hud = {"fmr_by_bedroom": {"two_br": 1_800}}  # only 2BR present
        result = self._mix().with_hud_fmr(hud)
        two_br = next(u for u in result["units"] if u["type_name"] == "2BR")
        assert two_br["hud_fmr"] == 1_800
        studio = next(u for u in result["units"] if u["type_name"] == "Studio")
        assert studio["hud_fmr"] is None


# ---------------------------------------------------------------------------
# HUD bedroom key mapping
# ---------------------------------------------------------------------------

class TestHudBedroomKeys:
    def test_all_unit_types_have_hud_key(self):
        for t in ("Studio", "1BR", "2BR", "3BR"):
            assert t in HUD_BEDROOM_KEYS


# ---------------------------------------------------------------------------
# parse_unit_mix_from_form
# ---------------------------------------------------------------------------

class TestParseUnitMixFromForm:
    def test_parses_valid_form(self):
        form = {
            "um_count_Studio": "10", "um_rent_Studio": "1100", "um_market_rent_Studio": "1200",
            "um_count_1BR": "30",    "um_rent_1BR": "1500",    "um_market_rent_1BR": "1600",
            "um_count_2BR": "50",    "um_rent_2BR": "1850",    "um_market_rent_2BR": "2000",
            "um_count_3BR": "10",    "um_rent_3BR": "2200",    "um_market_rent_3BR": "2400",
        }
        mix = parse_unit_mix_from_form(form)
        assert mix.total_units == 100
        assert not mix.is_empty()

    def test_skips_rows_with_zero_count(self):
        form = {
            "um_count_Studio": "0",  "um_rent_Studio": "1100",
            "um_count_1BR":    "10", "um_rent_1BR":    "1500",
        }
        mix = parse_unit_mix_from_form(form)
        assert mix.total_units == 10

    def test_skips_rows_with_zero_rent(self):
        form = {
            "um_count_Studio": "10", "um_rent_Studio": "0",
            "um_count_1BR":    "10", "um_rent_1BR":    "1500",
        }
        mix = parse_unit_mix_from_form(form)
        assert mix.total_units == 10

    def test_empty_form_returns_empty_mix(self):
        mix = parse_unit_mix_from_form({})
        assert mix.is_empty()

    def test_handles_invalid_strings_gracefully(self):
        form = {"um_count_1BR": "abc", "um_rent_1BR": "xyz"}
        mix = parse_unit_mix_from_form(form)
        assert mix.is_empty()
