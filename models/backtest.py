"""Historical backtest for rent prediction model.

Splits FRED CPI Shelter data into train/test by time, trains the predictor
on older data, compares predictions vs actuals. Returns MAE, RMSE,
direction accuracy, and chart data.
"""

import logging
import math
import numpy as np

logger = logging.getLogger(__name__)


class RentBacktester:
    """Backtest the rent prediction model using time-series split."""

    def __init__(self, train_pct: float = 0.70):
        self.train_pct = train_pct

    def run_backtest(self, cpi_shelter_data: list[dict],
                     zori_growth_rates: list[dict] = None) -> dict:
        """Run backtest on CPI Shelter data with time-based split.

        Args:
            cpi_shelter_data: List of dicts with 'date' and 'growth_rate' keys.
            zori_growth_rates: Optional ZORI growth rates for blending.

        Returns:
            Dict with MAE, RMSE, direction accuracy, actual vs predicted arrays.
        """
        if not cpi_shelter_data or len(cpi_shelter_data) < 10:
            return {"error": "Insufficient data for backtest (need >= 10 observations)"}

        try:
            from sklearn.preprocessing import PolynomialFeatures
            from sklearn.linear_model import LinearRegression

            rates = [d["growth_rate"] for d in cpi_shelter_data]
            dates = [d.get("date", f"Period {i}") for i, d in enumerate(cpi_shelter_data)]

            # Time-based split
            split_idx = int(len(rates) * self.train_pct)
            if split_idx < 5 or len(rates) - split_idx < 3:
                return {"error": "Insufficient data for meaningful split"}

            train_rates = rates[:split_idx]
            test_rates = rates[split_idx:]
            test_dates = dates[split_idx:]

            # Train polynomial regression on older data
            X_train = np.arange(len(train_rates)).reshape(-1, 1)
            y_train = np.array(train_rates)

            poly = PolynomialFeatures(degree=2, include_bias=False)
            X_poly = poly.fit_transform(X_train)

            model = LinearRegression()
            model.fit(X_poly, y_train)

            # Predict on test period
            test_steps = np.arange(split_idx, split_idx + len(test_rates)).reshape(-1, 1)
            X_test_poly = poly.transform(test_steps)
            predictions_raw = model.predict(X_test_poly)

            # Clamp predictions same as RentPredictor
            predictions = [max(-2.0, min(8.0, p)) for p in predictions_raw]

            # If ZORI available, also backtest blended
            blended_predictions = None
            if zori_growth_rates and len(zori_growth_rates) >= 2:
                zori_rates_vals = [r["growth_rate"] for r in zori_growth_rates]
                zori_avg = sum(zori_rates_vals) / len(zori_rates_vals)
                recent_zori = zori_rates_vals[-min(3, len(zori_rates_vals)):]
                zori_trend = sum(recent_zori) / len(recent_zori)

                blended = []
                for i, cpi_pred in enumerate(predictions):
                    zori_forecast = zori_trend * (0.9 ** i) + zori_avg * (1 - 0.9 ** i)
                    blend = 0.60 * zori_forecast + 0.40 * cpi_pred
                    blend = max(-2.0, min(8.0, blend))
                    blended.append(blend)
                blended_predictions = [round(b, 2) for b in blended]

            # Calculate metrics
            actuals = np.array(test_rates)
            preds = np.array(predictions)

            mae = float(np.mean(np.abs(actuals - preds)))
            rmse = float(math.sqrt(np.mean((actuals - preds) ** 2)))

            # Direction accuracy: did we predict the right direction of change?
            if len(actuals) >= 2:
                actual_dirs = np.sign(np.diff(actuals))
                pred_dirs = np.sign(np.diff(preds))
                direction_accuracy = float(np.mean(actual_dirs == pred_dirs) * 100)
            else:
                direction_accuracy = None

            # Mean prediction vs mean actual
            mean_actual = float(np.mean(actuals))
            mean_predicted = float(np.mean(preds))

            # Blended metrics
            blended_mae = None
            blended_rmse = None
            if blended_predictions:
                bp = np.array(blended_predictions)
                blended_mae = float(np.mean(np.abs(actuals - bp)))
                blended_rmse = float(math.sqrt(np.mean((actuals - bp) ** 2)))

            return {
                "train_periods": split_idx,
                "test_periods": len(test_rates),
                "total_periods": len(rates),
                "split_pct": round(self.train_pct * 100),
                # CPI-only model metrics
                "mae": round(mae, 3),
                "rmse": round(rmse, 3),
                "direction_accuracy": round(direction_accuracy, 1) if direction_accuracy is not None else None,
                "mean_actual": round(mean_actual, 2),
                "mean_predicted": round(mean_predicted, 2),
                # Blended model metrics (if ZORI available)
                "blended_mae": round(blended_mae, 3) if blended_mae else None,
                "blended_rmse": round(blended_rmse, 3) if blended_rmse else None,
                "has_blended": blended_predictions is not None,
                # Chart data
                "actual_values": [round(r, 2) for r in test_rates],
                "predicted_values": [round(p, 2) for p in predictions],
                "blended_values": blended_predictions,
                "test_dates": test_dates,
                "train_dates": dates[:split_idx],
                "train_values": [round(r, 2) for r in train_rates],
                # Quality assessment
                "quality": _assess_quality(mae, direction_accuracy),
            }

        except Exception as e:
            logger.error(f"Backtest failed: {e}")
            return {"error": str(e)}


def _assess_quality(mae: float, direction_accuracy: float | None) -> str:
    """Assess backtest quality based on MAE and direction accuracy."""
    if mae < 0.5 and direction_accuracy and direction_accuracy >= 70:
        return "STRONG"
    elif mae < 1.0 and (direction_accuracy is None or direction_accuracy >= 55):
        return "MODERATE"
    elif mae < 2.0:
        return "FAIR"
    else:
        return "WEAK"
