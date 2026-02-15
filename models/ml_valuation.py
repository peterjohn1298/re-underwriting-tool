"""ML-based property valuation using GradientBoosting regression.

IMPORTANT: This model trains on SYNTHETIC data calibrated to real macro indicators.
The R² score reflects fit on synthetic data, NOT predictive accuracy on real transactions.
Results should supplement — not replace — professional appraisals.
"""

import logging
import numpy as np
import pandas as pd
from sklearn.ensemble import GradientBoostingRegressor
from sklearn.model_selection import train_test_split

logger = logging.getLogger(__name__)


class PropertyValuationModel:
    """GradientBoosting property valuation model trained on synthetic data."""

    FEATURES = [
        "total_units", "sf_per_unit", "year_built", "occupancy",
        "in_place_rent", "market_rent", "noi_per_unit",
        "property_class_encoded", "market_cap_rate",
        "median_income", "population_millions", "unemployment_rate",
        "mortgage_rate",
        # Enhanced macro features
        "cpi_yoy_inflation", "housing_starts", "rental_vacancy_rate",
        "treasury_spread", "median_rent_census",
    ]

    def __init__(self):
        self.model = GradientBoostingRegressor(
            n_estimators=250,
            max_depth=4,
            learning_rate=0.1,
            min_samples_split=10,
            min_samples_leaf=5,
            random_state=42,
        )
        self.is_trained = False
        self.train_r2 = None
        self.test_r2 = None
        self.test_mae = None
        self.test_mape = None

    def _generate_training_data(self, market_data: dict) -> pd.DataFrame:
        """Generate synthetic training data calibrated to real market indicators."""
        np.random.seed(42)
        n_samples = 800

        # Extract real market calibration points
        demo = market_data.get("demographics", {}).get("structured", {})
        cap_data = market_data.get("cap_rates", {})
        macro = market_data.get("macro", {})

        real_median_income = demo.get("median_income") or 65000
        real_population = demo.get("population") or 5_000_000
        real_unemployment = demo.get("unemployment_rate") or 4.0
        real_cap_rate = cap_data.get("average_cap_rate") or 5.5
        treasury_10yr = cap_data.get("treasury_10yr")
        real_mortgage = (treasury_10yr + 1.75) if treasury_10yr else 6.75

        # New macro calibration points
        real_cpi_inflation = macro.get("cpi_yoy_inflation") or 3.2
        real_housing_starts = macro.get("housing_starts") or 1400
        real_rental_vacancy = macro.get("rental_vacancy_rate") or 6.5
        real_treasury_spread = macro.get("treasury_spread") or 0.5
        real_median_rent = demo.get("median_rent") or 1400

        records = []
        for i in range(n_samples):
            units = np.random.choice([20, 40, 60, 80, 100, 120, 150, 200, 250, 300])
            sf_per_unit = np.random.normal(850, 150)
            sf_per_unit = max(500, min(1500, sf_per_unit))
            year_built = int(np.random.normal(1995, 15))
            year_built = max(1960, min(2024, year_built))
            occupancy = np.random.normal(0.93, 0.04)
            occupancy = max(0.75, min(0.99, occupancy))

            base_rent = real_median_income / 12 * np.random.normal(0.35, 0.08)
            base_rent = max(600, min(3500, base_rent))
            in_place_rent = base_rent * np.random.normal(0.95, 0.05)
            market_rent = base_rent * np.random.normal(1.05, 0.05)

            expense_ratio = np.random.normal(0.42, 0.06)
            expense_ratio = max(0.30, min(0.55, expense_ratio))
            noi_per_unit = in_place_rent * 12 * occupancy * (1 - expense_ratio)

            prop_class = np.random.choice([1, 2, 3], p=[0.2, 0.5, 0.3])

            cap_rate = real_cap_rate + np.random.normal(0, 0.5)
            cap_rate = max(3.5, min(9.0, cap_rate))
            med_income = real_median_income * np.random.normal(1.0, 0.15)
            pop_m = (real_population / 1e6) * np.random.normal(1.0, 0.2)
            unemp = real_unemployment + np.random.normal(0, 0.8)
            unemp = max(2.0, min(10.0, unemp))
            mort_rate = real_mortgage + np.random.normal(0, 0.5)
            mort_rate = max(3.0, min(9.0, mort_rate))

            # New macro features with variation
            cpi_inf = real_cpi_inflation + np.random.normal(0, 1.0)
            cpi_inf = max(-1.0, min(10.0, cpi_inf))
            h_starts = real_housing_starts + np.random.normal(0, 200)
            h_starts = max(500, min(2500, h_starts))
            rent_vac = real_rental_vacancy + np.random.normal(0, 1.5)
            rent_vac = max(2.0, min(15.0, rent_vac))
            t_spread = real_treasury_spread + np.random.normal(0, 0.5)
            t_spread = max(-2.0, min(3.0, t_spread))
            med_rent_census = real_median_rent * np.random.normal(1.0, 0.15)
            med_rent_census = max(500, min(4000, med_rent_census))

            base_value_per_unit = (noi_per_unit / (cap_rate / 100))
            age_factor = 1.0 + (year_built - 1990) * 0.003
            class_factor = {1: 0.85, 2: 1.0, 3: 1.20}[prop_class]
            occ_factor = 1.0 + (occupancy - 0.93) * 2.0
            size_factor = 1.0 - (units - 100) * 0.0003
            income_factor = 1.0 + (med_income - 65000) / 200000
            # Macro adjustments to value
            inflation_factor = 1.0 + (cpi_inf - 3.0) * 0.01
            vacancy_factor = 1.0 - (rent_vac - 6.5) * 0.008
            starts_factor = 1.0 - (h_starts - 1400) * 0.00005

            value_per_unit = (base_value_per_unit * age_factor * class_factor
                              * occ_factor * size_factor * income_factor
                              * inflation_factor * vacancy_factor * starts_factor)
            noise = np.random.normal(1.0, 0.08)
            value_per_unit = max(50000, value_per_unit * noise)

            records.append({
                "total_units": units,
                "sf_per_unit": round(sf_per_unit),
                "year_built": year_built,
                "occupancy": round(occupancy, 3),
                "in_place_rent": round(in_place_rent),
                "market_rent": round(market_rent),
                "noi_per_unit": round(noi_per_unit),
                "property_class_encoded": prop_class,
                "market_cap_rate": round(cap_rate, 2),
                "median_income": round(med_income),
                "population_millions": round(pop_m, 2),
                "unemployment_rate": round(unemp, 1),
                "mortgage_rate": round(mort_rate, 2),
                "cpi_yoy_inflation": round(cpi_inf, 1),
                "housing_starts": round(h_starts),
                "rental_vacancy_rate": round(rent_vac, 1),
                "treasury_spread": round(t_spread, 2),
                "median_rent_census": round(med_rent_census),
                "value_per_unit": round(value_per_unit),
            })

        return pd.DataFrame(records)

    def train(self, market_data: dict):
        """Train on synthetic data with honest train/test split validation."""
        try:
            df = self._generate_training_data(market_data)
            X = df[self.FEATURES]
            y = df["value_per_unit"]

            # Honest train/test split (80/20)
            X_train, X_test, y_train, y_test = train_test_split(
                X, y, test_size=0.20, random_state=42
            )

            # Train on training set only
            self.model.fit(X_train, y_train)

            # Report metrics on HELD-OUT test set
            from sklearn.metrics import r2_score, mean_absolute_error
            y_pred_train = self.model.predict(X_train)
            y_pred_test = self.model.predict(X_test)

            self.train_r2 = round(float(r2_score(y_train, y_pred_train)), 4)
            self.test_r2 = round(float(r2_score(y_test, y_pred_test)), 4)
            self.test_mae = round(float(mean_absolute_error(y_test, y_pred_test)), 0)

            # MAPE on test set
            nonzero = y_test > 0
            if nonzero.sum() > 0:
                self.test_mape = round(float(
                    np.mean(np.abs((y_test[nonzero] - y_pred_test[nonzero]) / y_test[nonzero])) * 100
                ), 1)
            else:
                self.test_mape = None

            self.is_trained = True
            logger.info(
                f"ML Valuation trained ({len(self.FEATURES)} features, {len(X_train)} train / {len(X_test)} test). "
                f"Train R²={self.train_r2}, Test R²={self.test_r2}, "
                f"Test MAE=${self.test_mae:,.0f}, Test MAPE={self.test_mape}%"
            )

        except Exception as e:
            logger.error(f"ML model training failed: {e}")
            self.is_trained = False

    def predict(self, deal_inputs, derived, market_data: dict) -> dict:
        """Predict property value and return assessment with confidence context."""
        if not self.is_trained:
            return {"error": "Model not trained"}

        try:
            deal = deal_inputs if isinstance(deal_inputs, dict) else vars(deal_inputs)
            der = derived if isinstance(derived, dict) else vars(derived)

            demo = market_data.get("demographics", {}).get("structured", {})
            cap_data = market_data.get("cap_rates", {})
            macro = market_data.get("macro", {})

            class_str = deal.get("property_type", "")
            if "Class A" in class_str:
                prop_class = 3
            elif "Class C" in class_str:
                prop_class = 1
            else:
                prop_class = 2

            units = deal.get("total_units", 100)
            sf = deal.get("total_sf", 85000)
            noi = deal.get("current_noi", 0)

            features = {
                "total_units": units,
                "sf_per_unit": sf / units if units > 0 else 850,
                "year_built": deal.get("year_built", 2000),
                "occupancy": deal.get("occupancy", 0.92),
                "in_place_rent": deal.get("in_place_rent", 1300),
                "market_rent": deal.get("market_rent", 1450),
                "noi_per_unit": noi / units if units > 0 else 8000,
                "property_class_encoded": prop_class,
                "market_cap_rate": cap_data.get("average_cap_rate", 5.5),
                "median_income": demo.get("median_income") or 65000,
                "population_millions": (demo.get("population") or 5_000_000) / 1e6,
                "unemployment_rate": demo.get("unemployment_rate") or 4.0,
                "mortgage_rate": (cap_data.get("treasury_10yr", 4.5) + 1.75)
                                 if cap_data.get("treasury_10yr") else 6.75,
                # New macro features
                "cpi_yoy_inflation": macro.get("cpi_yoy_inflation") or 3.2,
                "housing_starts": macro.get("housing_starts") or 1400,
                "rental_vacancy_rate": macro.get("rental_vacancy_rate") or 6.5,
                "treasury_spread": macro.get("treasury_spread") or 0.5,
                "median_rent_census": demo.get("median_rent") or 1400,
            }

            X = pd.DataFrame([features])
            predicted_ppu = float(self.model.predict(X)[0])
            predicted_total = predicted_ppu * units

            actual_ppu = deal.get("purchase_price", 0) / units if units > 0 else 0
            premium_discount = ((actual_ppu - predicted_ppu) / predicted_ppu * 100
                                if predicted_ppu > 0 else 0)

            # Confidence-adjusted thresholds based on test MAPE
            threshold = max(10, (self.test_mape or 15) * 1.0)
            if premium_discount < -threshold:
                assessment = "UNDERVALUED"
            elif premium_discount > threshold:
                assessment = "OVERVALUED"
            else:
                assessment = "FAIR VALUE"

            # Feature importances
            importances = dict(zip(self.FEATURES,
                                   [round(float(x), 4) for x in self.model.feature_importances_]))
            sorted_imp = dict(sorted(importances.items(), key=lambda x: x[1], reverse=True))

            n_train = int(800 * 0.80)
            n_test = 800 - n_train

            return {
                "predicted_value_per_unit": round(predicted_ppu),
                "predicted_total_value": round(predicted_total),
                "actual_price_per_unit": round(actual_ppu),
                "premium_discount_pct": round(premium_discount, 1),
                "assessment": assessment,
                "feature_importances": sorted_imp,
                "train_r2": self.train_r2,
                "test_r2": self.test_r2,
                "r2_score": self.test_r2,  # backward compat
                "test_mae": self.test_mae,
                "test_mape": self.test_mape,
                "model_type": "GradientBoostingRegressor",
                "training_samples": n_train,
                "test_samples": n_test,
                "features_used": len(self.FEATURES),
                "confidence_note": (
                    f"Model trained on {800} synthetic records calibrated to FRED/Census macro indicators "
                    f"using {len(self.FEATURES)} features. "
                    f"Test set MAPE: {self.test_mape}%. Assessment threshold: +/-{threshold:.0f}%. "
                    f"This is NOT a substitute for a professional appraisal."
                ),
            }

        except Exception as e:
            logger.error(f"ML prediction failed: {e}")
            return {"error": str(e)}
