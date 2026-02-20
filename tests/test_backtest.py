"""Tests for RentBacktester."""
import pytest
from models.backtest import RentBacktester, _assess_quality


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _cpi_data(n=30):
    """Generate synthetic CPI Shelter data with a mild upward trend."""
    import math
    return [
        {"date": f"2020-{i:02d}", "growth_rate": 3.0 + 0.05 * i + 0.3 * math.sin(i)}
        for i in range(1, n + 1)
    ]


# ---------------------------------------------------------------------------
# Structure tests
# ---------------------------------------------------------------------------

class TestBacktestStructure:
    def test_returns_dict(self):
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30))
        assert isinstance(result, dict)

    def test_has_required_keys(self):
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30))
        for key in ("mae", "rmse", "train_periods", "test_periods",
                    "actual_values", "predicted_values", "quality"):
            assert key in result, f"Missing key: {key}"

    def test_train_test_split_adds_up(self):
        bt = RentBacktester(train_pct=0.70)
        result = bt.run_backtest(_cpi_data(30))
        assert result["train_periods"] + result["test_periods"] == result["total_periods"]

    def test_actual_and_predicted_same_length(self):
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30))
        assert len(result["actual_values"]) == len(result["predicted_values"])


# ---------------------------------------------------------------------------
# Metric validity
# ---------------------------------------------------------------------------

class TestBacktestMetrics:
    def test_mae_non_negative(self):
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30))
        assert result["mae"] >= 0

    def test_rmse_non_negative(self):
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30))
        assert result["rmse"] >= 0

    def test_rmse_ge_mae(self):
        # RMSE >= MAE is always true by Jensen's inequality
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30))
        assert result["rmse"] >= result["mae"] - 1e-9  # allow float tolerance

    def test_direction_accuracy_in_range(self):
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30))
        if result.get("direction_accuracy") is not None:
            assert 0 <= result["direction_accuracy"] <= 100

    def test_predictions_clamped_within_bounds(self):
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30))
        for p in result["predicted_values"]:
            assert -2.0 <= p <= 8.0


# ---------------------------------------------------------------------------
# Edge cases
# ---------------------------------------------------------------------------

class TestBacktestEdgeCases:
    def test_insufficient_data_returns_error(self):
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(5))
        assert "error" in result

    def test_empty_data_returns_error(self):
        bt = RentBacktester()
        result = bt.run_backtest([])
        assert "error" in result

    def test_none_data_returns_error(self):
        bt = RentBacktester()
        result = bt.run_backtest(None)
        assert "error" in result

    def test_custom_train_pct(self):
        bt = RentBacktester(train_pct=0.80)
        result = bt.run_backtest(_cpi_data(30))
        assert "error" not in result
        assert result["split_pct"] == 80


# ---------------------------------------------------------------------------
# Blended (ZORI) prediction
# ---------------------------------------------------------------------------

class TestBacktestBlended:
    def test_blended_present_when_zori_provided(self):
        zori = [{"growth_rate": 4.0 + i * 0.1} for i in range(10)]
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30), zori_growth_rates=zori)
        assert result["has_blended"] is True
        assert result["blended_values"] is not None

    def test_blended_mae_non_negative(self):
        zori = [{"growth_rate": 4.0} for _ in range(10)]
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30), zori_growth_rates=zori)
        if result.get("blended_mae") is not None:
            assert result["blended_mae"] >= 0

    def test_no_blended_without_zori(self):
        bt = RentBacktester()
        result = bt.run_backtest(_cpi_data(30))
        assert result["has_blended"] is False


# ---------------------------------------------------------------------------
# Quality assessment
# ---------------------------------------------------------------------------

class TestAssessQuality:
    def test_strong_quality(self):
        assert _assess_quality(mae=0.3, direction_accuracy=75) == "STRONG"

    def test_moderate_quality(self):
        assert _assess_quality(mae=0.7, direction_accuracy=60) == "MODERATE"

    def test_fair_quality(self):
        assert _assess_quality(mae=1.5, direction_accuracy=40) == "FAIR"

    def test_weak_quality(self):
        assert _assess_quality(mae=3.0, direction_accuracy=30) == "WEAK"

    def test_moderate_without_direction_accuracy(self):
        assert _assess_quality(mae=0.7, direction_accuracy=None) == "MODERATE"
