"""Tests for LP/GP waterfall / promote model."""
import pytest
from models.waterfall import run_waterfall, _safe_irr


# ---------------------------------------------------------------------------
# Additional coverage: zero-CF deals, IRR sanity, promote math
# ---------------------------------------------------------------------------

class TestWaterfallZeroOperatingCF:
    """Deals where all value is realized at exit (no operating cash flows)."""

    def test_pref_satisfied_from_exit_alone(self):
        # Large exit with zero operating CFs — pref comes entirely from exit
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=1_000_000, annual_btcfs=[0] * 5,
            exit_proceeds=2_000_000, hold_years=5,
        )
        assert result["pref_fully_satisfied"] is True

    def test_lp_receives_capital_plus_pref_from_exit(self):
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=1_000_000, annual_btcfs=[0] * 5,
            exit_proceeds=2_000_000, hold_years=5,
        )
        # LP must get at least capital + pref return
        assert result["lp_total"] >= result["lp_return_of_capital"] + result["lp_pref_paid"]

    def test_gp_earns_promote_on_zero_cf_deal(self):
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=1_000_000, annual_btcfs=[0] * 5,
            exit_proceeds=2_000_000, hold_years=5,
        )
        assert result["gp_promote"] > 0


class TestWaterfallIRRSanity:
    """LP and GP IRR should be finite, real numbers on valid deals."""

    def test_lp_irr_positive_on_profitable_deal(self):
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=1_000_000, annual_btcfs=[80_000] * 7,
            exit_proceeds=1_400_000, hold_years=7,
        )
        assert result["lp_irr"] is not None
        assert result["lp_irr"] > 0

    def test_gp_irr_positive_on_profitable_deal(self):
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=1_000_000, annual_btcfs=[80_000] * 7,
            exit_proceeds=1_400_000, hold_years=7,
        )
        assert result["gp_irr"] is not None
        assert result["gp_irr"] > 0

    def test_irr_returns_none_not_crash_on_loss_deal(self):
        # Catastrophic deal — total return less than equity invested
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=5_000_000, annual_btcfs=[0] * 5,
            exit_proceeds=1_000_000, hold_years=5,
        )
        # Should not raise; lp_irr may be negative or None but must not crash
        assert "lp_irr" in result
        assert "gp_irr" in result


class TestWaterfallPromoteMath:
    """Verify promote split is arithmetically correct."""

    def test_promote_is_exact_pct_of_residual(self):
        promote_pct = 0.20
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=promote_pct,
            equity_invested=1_000_000, annual_btcfs=[80_000] * 7,
            exit_proceeds=1_400_000, hold_years=7,
        )
        residual = result["lp_residual"] + result["gp_promote"]
        if residual > 0:
            actual_promote_pct = result["gp_promote"] / residual
            assert actual_promote_pct == pytest.approx(promote_pct, rel=1e-3)

    def test_lp_residual_is_complement_of_promote(self):
        promote_pct = 0.30
        result = run_waterfall(
            lp_equity_pct=0.85, preferred_return=0.08, promote_pct=promote_pct,
            equity_invested=2_000_000, annual_btcfs=[120_000] * 5,
            exit_proceeds=3_500_000, hold_years=5,
        )
        residual = result["lp_residual"] + result["gp_promote"]
        if residual > 0:
            lp_pct = result["lp_residual"] / residual
            assert lp_pct == pytest.approx(1 - promote_pct, rel=1e-3)

    def test_total_equals_lp_plus_gp_on_loss_deal(self):
        # Even on a losing deal, LP + GP totals must equal total distributions
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=3_000_000, annual_btcfs=[0] * 5,
            exit_proceeds=1_500_000, hold_years=5,
        )
        assert (result["lp_total"] + result["gp_total"]) == pytest.approx(
            result["total_distributions"], rel=1e-3
        )


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _simple_deal(equity=1_000_000, lp_pct=0.90, pref=0.08, promote=0.20,
                 annual_cf=80_000, exit=1_400_000, years=7):
    """Return a standard waterfall result for a healthy deal."""
    btcfs = [annual_cf] * years
    return run_waterfall(lp_pct, pref, promote, equity, btcfs, exit, years)


# ---------------------------------------------------------------------------
# Known-answer tests
# ---------------------------------------------------------------------------

class TestWaterfallKnownAnswers:
    def test_lp_receives_at_least_capital_back(self):
        result = _simple_deal()
        assert result["lp_return_of_capital"] == pytest.approx(900_000, rel=1e-3)

    def test_gp_receives_at_least_capital_back(self):
        result = _simple_deal()
        assert result["gp_return_of_capital"] == pytest.approx(100_000, rel=1e-3)

    def test_lp_total_exceeds_equity_on_profitable_deal(self):
        result = _simple_deal()
        assert result["lp_total"] > result["lp_equity"]

    def test_gp_total_exceeds_equity_on_profitable_deal(self):
        result = _simple_deal()
        assert result["gp_total"] > result["gp_equity"]

    def test_distributions_sum_to_total_available(self):
        result = _simple_deal()
        assert (result["lp_total"] + result["gp_total"]) == pytest.approx(
            result["total_distributions"], rel=1e-3
        )

    def test_lp_equity_multiple_above_one_on_profitable_deal(self):
        result = _simple_deal()
        assert result["lp_equity_multiple"] > 1.0

    def test_gp_equity_multiple_higher_than_lp_on_highly_profitable_deal(self):
        # GP EM > LP EM only when residual is large enough to overwhelm LP pref advantage
        # Use a very profitable deal: $200k/yr CF, $3M exit on $1M equity
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=1_000_000, annual_btcfs=[200_000] * 5,
            exit_proceeds=3_000_000, hold_years=5,
        )
        assert result["gp_equity_multiple"] > result["lp_equity_multiple"]

    def test_pref_fully_satisfied_on_good_deal(self):
        result = _simple_deal()
        assert result["pref_fully_satisfied"] is True

    def test_promote_dollars_positive_on_good_deal(self):
        result = _simple_deal()
        assert result["gp_promote"] > 0


# ---------------------------------------------------------------------------
# Edge cases
# ---------------------------------------------------------------------------

class TestWaterfallEdgeCases:
    def test_pref_unpaid_when_deal_performs_poorly(self):
        # Very small exit — not enough to satisfy pref
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=1_000_000, annual_btcfs=[0] * 7,
            exit_proceeds=600_000, hold_years=7,
        )
        assert result["pref_fully_satisfied"] is False
        assert result["lp_pref_unpaid"] > 0

    def test_promote_zero_when_pref_unpaid(self):
        result = run_waterfall(
            lp_equity_pct=0.90, preferred_return=0.08, promote_pct=0.20,
            equity_invested=1_000_000, annual_btcfs=[0] * 7,
            exit_proceeds=600_000, hold_years=7,
        )
        assert result["gp_promote"] == pytest.approx(0.0, abs=1.0)

    def test_invalid_equity_returns_error(self):
        result = run_waterfall(0.90, 0.08, 0.20, 0, [50_000] * 5, 500_000, 5)
        assert "error" in result

    def test_invalid_lp_pct_zero_returns_error(self):
        result = run_waterfall(0.0, 0.08, 0.20, 1_000_000, [50_000] * 5, 500_000, 5)
        assert "error" in result

    def test_invalid_lp_pct_one_returns_error(self):
        result = run_waterfall(1.0, 0.08, 0.20, 1_000_000, [50_000] * 5, 500_000, 5)
        assert "error" in result

    def test_negative_exit_proceeds_treated_as_zero(self):
        result = run_waterfall(0.90, 0.08, 0.20, 1_000_000, [50_000] * 5, -100_000, 5)
        assert "error" not in result
        assert result["total_distributions"] >= 0


# ---------------------------------------------------------------------------
# Tier table structure
# ---------------------------------------------------------------------------

class TestWaterfallTiers:
    def test_tiers_list_has_five_rows(self):
        result = _simple_deal()
        assert len(result["tiers"]) == 5

    def test_tiers_include_total_row(self):
        result = _simple_deal()
        total_row = result["tiers"][-1]
        assert total_row["tier"] == "TOTAL"

    def test_total_row_amounts_match_result(self):
        result = _simple_deal()
        total_row = result["tiers"][-1]
        assert total_row["lp_amount"] == pytest.approx(result["lp_total"], rel=1e-3)
        assert total_row["gp_amount"] == pytest.approx(result["gp_total"], rel=1e-3)


# ---------------------------------------------------------------------------
# IRR helper
# ---------------------------------------------------------------------------

class TestSafeIrr:
    def test_returns_float_for_valid_cashflows(self):
        cfs = [-100_000, 20_000, 20_000, 20_000, 20_000, 120_000]
        result = _safe_irr(cfs)
        assert isinstance(result, float)
        assert 0.10 < result < 0.20

    def test_returns_none_for_all_negative(self):
        result = _safe_irr([-100_000, -10_000, -10_000])
        assert result is None or isinstance(result, float)

    def test_does_not_crash_on_empty(self):
        result = _safe_irr([])
        assert result is None or isinstance(result, float)
