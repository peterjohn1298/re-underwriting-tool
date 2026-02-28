"""Tests for Monte Carlo simulation."""
import pytest
from models.monte_carlo import MonteCarloSimulator
from models.assumptions import DealInputs
from models.financial_model import build_pro_forma


def _make_deal(**overrides):
    defaults = dict(
        property_type="Multifamily - Class B",
        address="123 Main St, Austin, TX 78701",
        purchase_price=15_000_000,
        total_units=100,
        in_place_rent=1_500,
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


class TestMonteCarloSimulator:
    def test_returns_dict(self):
        sim = MonteCarloSimulator(n_iterations=100, seed=42)
        result = sim.run(_make_deal(), build_pro_forma)
        assert isinstance(result, dict)

    def test_has_irr_fields(self):
        sim = MonteCarloSimulator(n_iterations=100, seed=42)
        result = sim.run(_make_deal(), build_pro_forma)
        # Should have at least one IRR-related key
        irr_keys = [k for k in result if "irr" in k.lower()]
        assert len(irr_keys) > 0, f"No IRR keys found. Got: {list(result.keys())}"

    def test_percentiles_ordered(self):
        sim = MonteCarloSimulator(n_iterations=200, seed=42)
        result = sim.run(_make_deal(), build_pro_forma)
        p10 = result.get("irr_p10") or result.get("p10_irr")
        p50 = result.get("irr_p50") or result.get("p50_irr") or result.get("irr_median")
        p90 = result.get("irr_p90") or result.get("p90_irr")
        if p10 is not None and p50 is not None and p90 is not None:
            assert p10 <= p50 <= p90

    def test_probabilities_present_and_valid(self):
        sim = MonteCarloSimulator(n_iterations=100, seed=42)
        result = sim.run(_make_deal(), build_pro_forma)
        # Probabilities stored as percentages (0-100) under 'probabilities' key
        probs = result.get("probabilities", {})
        assert isinstance(probs, dict)
        assert len(probs) > 0
        for label, pct in probs.items():
            assert 0.0 <= pct <= 100.0, f"{label}: {pct} not in [0,100]"

    def test_different_seeds_give_different_results(self):
        sim1 = MonteCarloSimulator(n_iterations=50, seed=1)
        sim2 = MonteCarloSimulator(n_iterations=50, seed=999)
        deal = _make_deal()
        r1 = sim1.run(deal, build_pro_forma)
        r2 = sim2.run(deal, build_pro_forma)
        # Results should differ with different seeds
        irr_key = next((k for k in r1 if "irr" in k.lower() and "mean" in k.lower()), None)
        if irr_key and irr_key in r2:
            assert r1[irr_key] != r2[irr_key]

    def test_does_not_crash_with_minimal_deal(self):
        sim = MonteCarloSimulator(n_iterations=50, seed=42)
        result = sim.run(_make_deal(purchase_price=1_000_000, total_units=10), build_pro_forma)
        assert result is not None


class TestMonteCarloCholesky:
    """Verify Cholesky-correlated shock metadata is present and valid."""

    def test_shock_method_is_cholesky(self):
        sim = MonteCarloSimulator(n_iterations=100, seed=42)
        result = sim.run(_make_deal(), build_pro_forma)
        assert result.get("shock_method") == "cholesky_correlated_normal"

    def test_correlation_matrix_in_result(self):
        sim = MonteCarloSimulator(n_iterations=100, seed=42)
        result = sim.run(_make_deal(), build_pro_forma)
        corr = result.get("correlation_matrix")
        assert corr is not None
        assert "variables" in corr
        assert "matrix" in corr
        assert len(corr["variables"]) == 4
        assert len(corr["matrix"]) == 4

    def test_correlation_matrix_is_symmetric(self):
        sim = MonteCarloSimulator(n_iterations=50, seed=42)
        result = sim.run(_make_deal(), build_pro_forma)
        matrix = result["correlation_matrix"]["matrix"]
        for i in range(4):
            for j in range(4):
                assert matrix[i][j] == pytest.approx(matrix[j][i], abs=1e-9)

    def test_correlation_diagonal_is_one(self):
        sim = MonteCarloSimulator(n_iterations=50, seed=42)
        result = sim.run(_make_deal(), build_pro_forma)
        matrix = result["correlation_matrix"]["matrix"]
        for i in range(4):
            assert matrix[i][i] == pytest.approx(1.0, abs=1e-9)

    def test_different_seeds_produce_different_distributions(self):
        deal = _make_deal()
        r1 = MonteCarloSimulator(n_iterations=200, seed=7).run(deal, build_pro_forma)
        r2 = MonteCarloSimulator(n_iterations=200, seed=999).run(deal, build_pro_forma)
        assert r1.get("mean_irr") != r2.get("mean_irr")
