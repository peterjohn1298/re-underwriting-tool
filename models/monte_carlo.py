"""Monte Carlo simulation for real estate investment returns.

Randomizes key assumptions (rent growth, occupancy, exit cap, expense growth)
and runs build_pro_forma() N times to produce an IRR distribution.
"""

import logging
import copy
import numpy as np

logger = logging.getLogger(__name__)


class MonteCarloSimulator:
    """Monte Carlo simulation engine for RE underwriting."""

    def __init__(self, n_iterations: int = 1000, seed: int = 42):
        self.n_iterations = n_iterations
        self.seed = seed

    def run(self, deal, build_pro_forma_fn) -> dict:
        """Run Monte Carlo simulation.

        Args:
            deal: DealInputs dataclass instance
            build_pro_forma_fn: The build_pro_forma function to call

        Returns:
            Dict with IRR distribution, percentiles, probabilities,
            histogram data, and summary statement.
        """
        try:
            np.random.seed(self.seed)
            irrs = []
            equity_multiples = []
            failed = 0

            base_growth = deal.revenue_growth_rate
            base_occ = deal.occupancy
            base_exit_spread = deal.exit_cap_rate_spread
            base_expense_growth = deal.expense_growth_rate

            for i in range(self.n_iterations):
                sim_deal = copy.deepcopy(deal)

                # Randomize rent growth: ±30% of base
                growth_shock = np.random.uniform(-0.30, 0.30)
                sim_deal.revenue_growth_rate = max(0.0, base_growth * (1 + growth_shock))
                # Clear variable growth so we use flat rate
                sim_deal.yearly_revenue_growth = []

                # Randomize occupancy: ±3 percentage points
                occ_shock = np.random.uniform(-0.03, 0.03)
                sim_deal.occupancy = max(0.60, min(0.99, base_occ + occ_shock))

                # Randomize exit cap spread: ±50bps
                cap_shock = np.random.uniform(-0.0050, 0.0050)
                sim_deal.exit_cap_rate_spread = max(0.0, base_exit_spread + cap_shock)

                # Randomize expense growth: ±25% of base
                exp_shock = np.random.uniform(-0.25, 0.25)
                sim_deal.expense_growth_rate = max(0.0, base_expense_growth * (1 + exp_shock))

                try:
                    result = build_pro_forma_fn(sim_deal)
                    irr = result["metrics"].get("levered_irr")
                    em = result["metrics"].get("equity_multiple")
                    if irr is not None and -1.0 <= irr <= 2.0:
                        irrs.append(irr * 100)  # Convert to percentage
                        equity_multiples.append(em if em else 0)
                    else:
                        failed += 1
                except Exception:
                    failed += 1

            if len(irrs) < 10:
                return {"error": f"Too few valid simulations ({len(irrs)}/{self.n_iterations})"}

            irrs = np.array(irrs)
            ems = np.array(equity_multiples)

            # Percentiles
            percentiles = {
                "P5": round(float(np.percentile(irrs, 5)), 1),
                "P10": round(float(np.percentile(irrs, 10)), 1),
                "P25": round(float(np.percentile(irrs, 25)), 1),
                "P50 (Median)": round(float(np.percentile(irrs, 50)), 1),
                "P75": round(float(np.percentile(irrs, 75)), 1),
                "P90": round(float(np.percentile(irrs, 90)), 1),
                "P95": round(float(np.percentile(irrs, 95)), 1),
            }

            # Probabilities
            prob_8 = round(float(np.mean(irrs > 8) * 100), 1)
            prob_12 = round(float(np.mean(irrs > 12) * 100), 1)
            prob_15 = round(float(np.mean(irrs > 15) * 100), 1)
            prob_negative = round(float(np.mean(irrs < 0) * 100), 1)

            probabilities = {
                "IRR > 8%": prob_8,
                "IRR > 12%": prob_12,
                "IRR > 15%": prob_15,
                "IRR < 0% (Loss)": prob_negative,
            }

            # Histogram
            n_bins = 20
            hist_counts, hist_edges = np.histogram(irrs, bins=n_bins)
            histogram_bins = [round(float(e), 1) for e in hist_edges]
            histogram_counts = [int(c) for c in hist_counts]

            # Summary
            median_irr = percentiles["P50 (Median)"]
            summary = (
                f"{prob_12}% probability of IRR > 12%. "
                f"Median IRR: {median_irr:.1f}%. "
                f"Range: {percentiles['P5']:.1f}% (P5) to {percentiles['P95']:.1f}% (P95)."
            )

            # MC signal for recommendation
            if prob_12 >= 70:
                mc_signal = "POSITIVE"
                mc_detail = f"{prob_12}% chance of exceeding 12% IRR"
            elif prob_8 >= 70:
                mc_signal = "MODERATE"
                mc_detail = f"{prob_8}% chance of exceeding 8% IRR"
            elif prob_negative > 20:
                mc_signal = "WARNING"
                mc_detail = f"{prob_negative}% chance of negative returns"
            else:
                mc_signal = "NEUTRAL"
                mc_detail = f"Median IRR: {median_irr:.1f}%"

            return {
                "n_iterations": self.n_iterations,
                "valid_iterations": len(irrs),
                "failed_iterations": failed,
                "mean_irr": round(float(np.mean(irrs)), 1),
                "std_irr": round(float(np.std(irrs)), 1),
                "min_irr": round(float(np.min(irrs)), 1),
                "max_irr": round(float(np.max(irrs)), 1),
                "mean_em": round(float(np.mean(ems)), 2),
                "percentiles": percentiles,
                "probabilities": probabilities,
                "histogram_bins": histogram_bins,
                "histogram_counts": histogram_counts,
                "summary": summary,
                "mc_signal": mc_signal,
                "mc_detail": mc_detail,
            }

        except Exception as e:
            logger.error(f"Monte Carlo simulation failed: {e}")
            return {"error": str(e)}
