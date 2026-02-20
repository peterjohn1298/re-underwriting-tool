"""Tests for Base / Bull / Bear scenario comparison."""
import pytest
from models.scenarios import run_scenarios, BULL_OFFSETS, BEAR_OFFSETS
from models.assumptions import DealInputs
from models.financial_model import build_pro_forma


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _standard_deal():
    """100-unit multifamily deal at $15M — healthy base case."""
    return DealInputs(
        property_type="Multifamily - Class B",
        purchase_price=15_000_000,
        current_noi=795_000,
        total_units=100,
        in_place_rent=1_850,
        market_rent=2_000,
        occupancy=0.93,
        hold_period_years=7,
        ltv=0.65,
        interest_rate=0.065,
        revenue_growth_rate=0.03,
        expense_growth_rate=0.03,
        exit_cap_rate_spread=0.0025,
    )


# ---------------------------------------------------------------------------
# Structure tests
# ---------------------------------------------------------------------------

class TestScenariosStructure:
    def test_returns_base_bull_bear_keys(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        assert "base" in result
        assert "bull" in result
        assert "bear" in result

    def test_returns_comparison_table(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        assert "comparison_table" in result
        assert len(result["comparison_table"]) > 0

    def test_comparison_table_has_metric_base_bull_bear_columns(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        row = result["comparison_table"][0]
        assert "metric" in row
        assert "base" in row
        assert "bull" in row
        assert "bear" in row

    def test_each_scenario_has_levered_irr(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        for label in ("base", "bull", "bear"):
            assert "levered_irr" in result[label]

    def test_offsets_stored_in_result(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        assert "bull_offsets" in result
        assert "bear_offsets" in result


# ---------------------------------------------------------------------------
# Ordering / reasonableness
# ---------------------------------------------------------------------------

class TestScenariosOrdering:
    def test_bull_irr_exceeds_base_irr(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        assert result["bull"]["levered_irr"] > result["base"]["levered_irr"]

    def test_base_irr_exceeds_bear_irr(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        assert result["base"]["levered_irr"] > result["bear"]["levered_irr"]

    def test_bull_equity_multiple_exceeds_bear(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        assert result["bull"]["equity_multiple"] > result["bear"]["equity_multiple"]

    def test_base_occupancy_matches_deal(self):
        deal = _standard_deal()
        result = run_scenarios(deal, build_pro_forma)
        assert result["base"]["occupancy"] == pytest.approx(deal.occupancy, rel=1e-3)

    def test_bull_occupancy_higher_than_base(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        assert result["bull"]["occupancy"] > result["base"]["occupancy"]

    def test_bear_occupancy_lower_than_base(self):
        result = run_scenarios(_standard_deal(), build_pro_forma)
        assert result["bear"]["occupancy"] < result["base"]["occupancy"]


# ---------------------------------------------------------------------------
# Edge cases
# ---------------------------------------------------------------------------

class TestScenariosEdgeCases:
    def test_occupancy_clamped_at_minimum(self):
        deal = _standard_deal()
        deal.occupancy = 0.52  # Close to 0.50, bear will try to go below
        result = run_scenarios(deal, build_pro_forma,
                               bear_offsets={"occupancy": -0.10})
        assert result["bear"]["occupancy"] >= 0.50

    def test_occupancy_clamped_at_maximum(self):
        deal = _standard_deal()
        deal.occupancy = 0.98
        result = run_scenarios(deal, build_pro_forma,
                               bull_offsets={"occupancy": +0.10})
        assert result["bull"]["occupancy"] <= 1.0

    def test_custom_offsets_override_defaults(self):
        custom_bull = {"revenue_growth_rate": +0.05}
        result = run_scenarios(_standard_deal(), build_pro_forma,
                               bull_offsets=custom_bull)
        assert result["bull_offsets"] == custom_bull

    def test_no_crash_on_minimal_deal(self):
        deal = DealInputs(
            purchase_price=1_000_000,
            current_noi=60_000,
            total_units=10,
            in_place_rent=1_000,
            hold_period_years=5,
        )
        result = run_scenarios(deal, build_pro_forma)
        assert "base" in result


# ---------------------------------------------------------------------------
# Default offsets
# ---------------------------------------------------------------------------

class TestDefaultOffsets:
    def test_bull_offsets_have_positive_revenue_growth(self):
        assert BULL_OFFSETS["revenue_growth_rate"] > 0

    def test_bear_offsets_have_negative_revenue_growth(self):
        assert BEAR_OFFSETS["revenue_growth_rate"] < 0

    def test_bull_offsets_have_positive_occupancy(self):
        assert BULL_OFFSETS["occupancy"] > 0

    def test_bear_offsets_have_negative_occupancy(self):
        assert BEAR_OFFSETS["occupancy"] < 0
