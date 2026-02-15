"""Predictive rent growth model using polynomial regression on FRED CPI Shelter data.

Optionally blends Zillow ZORI city-level growth rates when available
(60% ZORI + 40% CPI for city-specific accuracy).
"""

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
        self.zori_growth_rates = None
        self.blend_used = False

    def train(self, cpi_shelter_data: list[dict], zori_growth_rates: list[dict] = None):
        """Train on FRED CPI Shelter annual growth rate data.

        Args:
            cpi_shelter_data: List of dicts with 'date' and 'growth_rate' keys,
                              from FREDClient.get_cpi_shelter().
            zori_growth_rates: Optional list of dicts with 'year' and 'growth_rate' keys,
                               from ZillowClient.get_annual_growth_rates().
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

            # Store ZORI data for blending during prediction
            if zori_growth_rates and len(zori_growth_rates) >= 2:
                self.zori_growth_rates = zori_growth_rates
                logger.info(f"Rent predictor trained on {len(rates)} CPI points + {len(zori_growth_rates)} ZORI points")
            else:
                self.zori_growth_rates = None
                logger.info(f"Rent predictor trained on {len(rates)} CPI data points")

        except Exception as e:
            logger.error(f"Rent predictor training failed: {e}")
            self.is_trained = False

    def predict(self, hold_period: int, current_rent: float,
                total_units: int) -> dict:
        """Forecast rent growth rates for the hold period.

        When ZORI data is available, blends 60% ZORI trend + 40% CPI model
        for city-specific accuracy.

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

            # CPI model predictions
            future_steps = np.arange(n_hist, n_hist + hold_period).reshape(-1, 1)
            X_future = self.poly.transform(future_steps)
            cpi_predictions = self.model.predict(X_future)

            # Blend with ZORI if available
            self.blend_used = False
            if self.zori_growth_rates and len(self.zori_growth_rates) >= 2:
                zori_rates = [r["growth_rate"] for r in self.zori_growth_rates]
                zori_avg = sum(zori_rates) / len(zori_rates)
                # Recent trend from last 2-3 ZORI data points
                recent_zori = zori_rates[-min(3, len(zori_rates)):]
                zori_trend = sum(recent_zori) / len(recent_zori)

                blended = []
                for i, cpi_rate in enumerate(cpi_predictions):
                    # ZORI forecast: slight mean-reversion toward average
                    zori_forecast = zori_trend * (0.9 ** i) + zori_avg * (1 - 0.9 ** i)
                    # 60% ZORI + 40% CPI
                    blend = 0.60 * zori_forecast + 0.40 * cpi_rate
                    blended.append(blend)

                raw_predictions = blended
                self.blend_used = True
                method_note = f"Blended: 60% Zillow ZORI + 40% CPI Shelter (degree={self.degree})"
                data_source = "Zillow ZORI + FRED CPI Shelter (CUSR0000SAH1)"
            else:
                raw_predictions = cpi_predictions
                method_note = f"Polynomial Regression (degree={self.degree})"
                data_source = "FRED CPI Shelter (CUSR0000SAH1)"

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

            result = {
                "method": method_note,
                "historical_rates": [round(r, 2) for r in recent_hist],
                "historical_avg": round(sum(recent_hist) / len(recent_hist), 2),
                "predicted_rates": predicted_rates,
                "predicted_rents_per_unit": predicted_rents,
                "predicted_annual_revenue": [round(r * total_units * 12) for r in predicted_rents],
                "avg_predicted_growth": avg_growth,
                "hold_period": hold_period,
                "current_rent": current_rent,
                "total_units": total_units,
                "data_source": data_source,
                "training_points": n_hist,
                "zori_blended": self.blend_used,
            }

            if self.blend_used and self.zori_growth_rates:
                result["zori_city"] = self.zori_growth_rates[0].get("year", "")
                result["zori_data_points"] = len(self.zori_growth_rates)
                result["zori_avg_growth"] = round(
                    sum(r["growth_rate"] for r in self.zori_growth_rates) / len(self.zori_growth_rates), 2
                )

            return result

        except Exception as e:
            logger.error(f"Rent prediction failed: {e}")
            return {"error": str(e)}
