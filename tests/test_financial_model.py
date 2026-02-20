"""Tests for the 10-year pro forma financial model."""
import pytest
from models.assumptions import DealInputs
from models.financial_model import build_pro_forma


def _make_deal(**overrides):
    """Build a valid DealInputs for testing. current_noi drives expense model."""
    defaults = dict(
        property_type="Multifamily - Class B",
        address="123 Main St, Austin, TX 78701",
        purchase_price=15_000_000,
        total_units=100,
        in_place_rent=1_500,
        current_noi=800_000,      # explicit NOI so expenses are non-trivial
        occupancy=0.93,
        hold_period_years=5,
        ltv=0.70,
        interest_rate=0.065,
        revenue_growth_rate=0.03,
        expense_growth_rate=0.025,
        enable_ml_valuation=False,
        enable_rent_prediction=False,
    )
    defaults.update(overrides)
    return DealInputs(**defaults)


class TestProFormaStructure:
    def test_returns_dict(self):
        result = build_pro_forma(_make_deal())
        assert isinstance(result, dict)

    def test_pro_forma_rows_present(self):
        result = build_pro_forma(_make_deal())
        assert "pro_forma" in result
        assert isinstance(result["pro_forma"], list)
        assert len(result["pro_forma"]) > 0

    def test_pro_forma_has_noi(self):
        result = build_pro_forma(_make_deal())
        for row in result["pro_forma"]:
            assert "noi" in row

    def test_noi_is_positive(self):
        result = build_pro_forma(_make_deal())
        for row in result["pro_forma"]:
            assert row["noi"] > 0, f"NOI should be positive, got {row['noi']}"

    def test_revenue_grows_over_time(self):
        result = build_pro_forma(_make_deal(revenue_growth_rate=0.03))
        egi = [row["effective_gross_income"] for row in result["pro_forma"]]
        for i in range(1, len(egi)):
            assert egi[i] >= egi[i - 1], "Revenue should grow year over year"

    def test_exit_value_present_and_positive(self):
        result = build_pro_forma(_make_deal())
        rev = result.get("reversion", {})
        sale_price = rev.get("sale_price", 0)
        assert sale_price > 0, f"Expected positive sale price, got {sale_price}"

    def test_loan_amount_matches_ltv(self):
        result = build_pro_forma(_make_deal(purchase_price=10_000_000, ltv=0.65))
        sources = result.get("sources_uses", {})
        loan = sources.get("loan_amount", result.get("loan_amount"))
        if loan is not None:
            assert loan == pytest.approx(6_500_000, rel=0.01)

    def test_metrics_present(self):
        result = build_pro_forma(_make_deal())
        assert "metrics" in result
        assert isinstance(result["metrics"], dict)

    def test_irr_field_exists(self):
        result = build_pro_forma(_make_deal())
        metrics = result.get("metrics", {})
        assert "levered_irr" in metrics

    def test_dscr_yr1_exists(self):
        result = build_pro_forma(_make_deal())
        metrics = result.get("metrics", {})
        assert "dscr_yr1" in metrics

    def test_hold_period_reflected_in_exit_year(self):
        result = build_pro_forma(_make_deal(hold_period_years=7))
        rev = result.get("reversion", {})
        assert rev.get("exit_year") == 7


class TestEdgeCases:
    def test_high_ltv_does_not_crash(self):
        result = build_pro_forma(_make_deal(ltv=0.85))
        assert result is not None

    def test_low_occupancy_does_not_crash(self):
        result = build_pro_forma(_make_deal(occupancy=0.70))
        assert result is not None

    def test_zero_capex_does_not_crash(self):
        result = build_pro_forma(_make_deal(planned_capex=0))
        assert result is not None

    def test_large_deal_does_not_crash(self):
        result = build_pro_forma(_make_deal(
            purchase_price=100_000_000, total_units=500, current_noi=5_000_000
        ))
        assert result is not None
