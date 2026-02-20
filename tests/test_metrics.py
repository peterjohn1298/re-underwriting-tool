"""Tests for financial metrics calculations."""
import pytest
from models.metrics import (
    calc_irr, calc_equity_multiple, calc_dscr,
    calc_cash_on_cash, calc_yield_on_cost,
)


class TestIRR:
    def test_positive_irr(self):
        # -100k invested, +20k/yr for 5 yrs + 120k exit → ~12.6% IRR
        cash_flows = [-100_000, 20_000, 20_000, 20_000, 20_000, 120_000]
        irr = calc_irr(cash_flows)
        assert irr is not None
        assert 0.10 < irr < 0.20

    def test_negative_cashflows_returns_none_or_negative(self):
        cash_flows = [-100_000, 1_000, 1_000, 1_000]
        irr = calc_irr(cash_flows)
        assert irr is None or irr < 0

    def test_irr_with_single_cashflow(self):
        cash_flows = [-100_000]
        irr = calc_irr(cash_flows)
        assert irr is None or isinstance(irr, float)

    def test_zero_investment_does_not_crash(self):
        cash_flows = [0, 10_000, 10_000]
        irr = calc_irr(cash_flows)
        assert irr is None or isinstance(irr, float)


class TestEquityMultiple:
    def test_basic_2x_multiple(self):
        # -100k in, +200k total out
        cash_flows = [-100_000, 50_000, 50_000, 100_000]
        em = calc_equity_multiple(cash_flows)
        assert em == pytest.approx(2.0, rel=1e-3)

    def test_zero_equity_returns_zero(self):
        cash_flows = [0, 50_000, 50_000]
        result = calc_equity_multiple(cash_flows)
        assert result == 0.0

    def test_loss_scenario(self):
        cash_flows = [-100_000, 20_000, 20_000, 20_000]
        em = calc_equity_multiple(cash_flows)
        assert em < 1.0


class TestDSCR:
    def test_healthy_dscr(self):
        dscr = calc_dscr(noi=150_000, annual_debt_service=100_000)
        assert dscr == pytest.approx(1.5, rel=1e-3)

    def test_dscr_below_1(self):
        dscr = calc_dscr(noi=80_000, annual_debt_service=100_000)
        assert dscr == pytest.approx(0.8, rel=1e-3)

    def test_zero_debt_service_returns_zero(self):
        result = calc_dscr(noi=100_000, annual_debt_service=0)
        assert result == 0.0


class TestCashOnCash:
    def test_15_pct_coc(self):
        coc = calc_cash_on_cash(annual_cf=15_000, equity_invested=100_000)
        assert coc == pytest.approx(0.15, rel=1e-3)

    def test_zero_equity_returns_zero(self):
        result = calc_cash_on_cash(annual_cf=15_000, equity_invested=0)
        assert result == 0.0

    def test_negative_cashflow(self):
        coc = calc_cash_on_cash(annual_cf=-5_000, equity_invested=100_000)
        assert coc == pytest.approx(-0.05, rel=1e-3)


class TestYieldOnCost:
    def test_10_pct_yield(self):
        yoc = calc_yield_on_cost(noi=100_000, total_cost=1_000_000)
        assert yoc == pytest.approx(0.10, rel=1e-3)

    def test_zero_cost_returns_zero(self):
        result = calc_yield_on_cost(noi=100_000, total_cost=0)
        assert result == 0.0
