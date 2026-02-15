"""Predictive rent growth model using polynomial regression on FRED CPI Shelter data."""

import logging
import numpy as np

logger = logging.getLogger(__name__)


class RentPredictor:
    """Polynomial regression model for forecasting rent growth rates."""

    def __init__(self, degree: int = 2):
        self.degree = degree
        self.model = None
        self.poly = None
        self.is_trained = False

    def train(self, cpi_shelter_data: list[dict]):
        """Train on FRED CPI Shelter annual growth rate data.

        Args:
            cpi_shelter_data: List of dicts with 'date' and 'growth_rate' keys,
                              from FREDClient.get_cpi_shelter().
        """
        try:
            from sklearn.preprocessing import PolynomialFeatures
            from sklearn.linear_model import LinearRegression

            if not cpi_shelter_data or len(cpi_shelter_data) < 5:
                logger.warning("Insufficient CPI Shelter data for training")
                return

            # Use sequential indices as features (time steps)
            rates = [d["growth_rate"] for d in cpi_shelter_data]
            X = np.arange(len(rates)).reshape(-1, 1)
            y = np.array(rates)

            self.poly = PolynomialFeatures(degree=self.degree, include_bias=False)
            X_poly = self.poly.fit_transform(X)

            self.model = LinearRegression()
            self.model.fit(X_poly, y)

            self.historical_rates = rates
            self.is_trained = True
            logger.info(f"Rent predictor trained on {len(rates)} data points")

        except Exception as e:
            logger.error(f"Rent predictor training failed: {e}")
            self.is_trained = False

    def predict(self, hold_period: int, current_rent: float,
                total_units: int) -> dict:
        """Forecast rent growth rates for the hold period.

        Args:
            hold_period: Number of years to forecast
            current_rent: Current monthly rent per unit
            total_units: Number of units

        Returns:
            Dict with predicted rates, predicted rents, and metadata
        """
        if not self.is_trained:
            return {"error": "Model not trained"}

        try:
            n_hist = len(self.historical_rates)

            # Predict future time steps
            future_steps = np.arange(n_hist, n_hist + hold_period).reshape(-1, 1)
            X_future = self.poly.transform(future_steps)
            raw_predictions = self.model.predict(X_future)

            # Clamp to [-2%, +8%] to prevent unrealistic forecasts
            predicted_rates = [round(max(-2.0, min(8.0, r)), 2) for r in raw_predictions]

            # Project rents forward
            predicted_rents = []
            rent = current_rent
            for rate in predicted_rates:
                rent = rent * (1 + rate / 100)
                predicted_rents.append(round(rent, 2))

            avg_growth = round(sum(predicted_rates) / len(predicted_rates), 2)

            # Historical summary (last 5)
            recent_hist = self.historical_rates[-5:] if len(self.historical_rates) >= 5 else self.historical_rates

            return {
                "method": f"Polynomial Regression (degree={self.degree})",
                "historical_rates": [round(r, 2) for r in recent_hist],
                "historical_avg": round(sum(recent_hist) / len(recent_hist), 2),
                "predicted_rates": predicted_rates,
                "predicted_rents_per_unit": predicted_rents,
                "predicted_annual_revenue": [round(r * total_units * 12) for r in predicted_rents],
                "avg_predicted_growth": avg_growth,
                "hold_period": hold_period,
                "current_rent": current_rent,
                "total_units": total_units,
                "data_source": "FRED CPI Shelter (CUSR0000SAH1)",
                "training_points": n_hist,
            }

        except Exception as e:
            logger.error(f"Rent prediction failed: {e}")
            return {"error": str(e)}
