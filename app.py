"""Flask application for Real Estate Investment Underwriting Tool."""

import os
import uuid
import threading
import logging
from flask import Flask, render_template, request, jsonify, send_file

from config import OUTPUT_DIR, UPLOAD_DIR, PORT, DEBUG, SECRET_KEY
from models.assumptions import DealInputs, derive_assumptions
from models.financial_model import build_pro_forma
from services.market_research import run_full_research
from services.excel_generator import generate_excel
from services.word_generator import generate_word

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = SECRET_KEY

# In-memory job store
jobs: dict[str, dict] = {}


def _f(val, default=0.0):
    """Parse float from form."""
    try:
        return float(val) if val else default
    except (ValueError, TypeError):
        return default


def _i(val, default=0):
    """Parse int from form."""
    try:
        return int(val) if val else default
    except (ValueError, TypeError):
        return default


def _parse_form(form) -> DealInputs:
    """Parse form data into DealInputs."""
    return DealInputs(
        property_type=form.get("property_type", "Multifamily - Class B"),
        address=form.get("address", ""),
        year_built=_i(form.get("year_built"), 2000),
        purchase_price=_f(form.get("purchase_price")),
        current_noi=_f(form.get("current_noi")),
        total_units=_i(form.get("total_units")),
        total_sf=_f(form.get("total_sf")),
        in_place_rent=_f(form.get("in_place_rent")),
        market_rent=_f(form.get("market_rent")),
        occupancy=_f(form.get("occupancy"), 92) / 100,
        deferred_maintenance=_f(form.get("deferred_maintenance")),
        planned_capex=_f(form.get("planned_capex")),
        capex_description=form.get("capex_description", ""),
        hold_period_years=_i(form.get("hold_period_years"), 7),
        # Overrides
        ltv=_f(form.get("ltv"), 65) / 100,
        interest_rate=_f(form.get("interest_rate"), 6.75) / 100,
        amortization_years=_i(form.get("amortization_years"), 30),
        loan_term_years=_i(form.get("loan_term_years"), 10),
        io_period_years=_i(form.get("io_period_years"), 0),
        closing_costs_pct=_f(form.get("closing_costs_pct"), 3) / 100,
        revenue_growth_rate=_f(form.get("revenue_growth_rate"), 3) / 100,
        expense_growth_rate=_f(form.get("expense_growth_rate"), 3) / 100,
        management_fee_pct=_f(form.get("management_fee_pct"), 3.5) / 100,
        exit_cap_rate_spread=_f(form.get("exit_cap_rate_spread"), 25) / 10000,
        sale_costs_pct=_f(form.get("sale_costs_pct"), 2.5) / 100,
        replacement_reserves_per_unit=_f(form.get("replacement_reserves_per_unit"), 250),
        # Tax analysis (Fix #9)
        tax_rate=_f(form.get("tax_rate"), 25) / 100,
        land_value_pct=_f(form.get("land_value_pct"), 20) / 100,
        # Expense overrides (Fix #10)
        override_property_tax=_f(form.get("override_property_tax")),
        override_insurance=_f(form.get("override_insurance")),
        override_utilities=_f(form.get("override_utilities")),
        override_repairs=_f(form.get("override_repairs")),
        override_general_admin=_f(form.get("override_general_admin")),
        override_other_expenses=_f(form.get("override_other_expenses")),
        # AI/ML features
        enable_ml_valuation=form.get("enable_ml_valuation") == "on",
        enable_rent_prediction=form.get("enable_rent_prediction") == "on",
    )


def _run_analysis(job_id: str, deal: DealInputs):
    """Run the full analysis pipeline in a background thread."""
    try:
        derived = derive_assumptions(deal)

        jobs[job_id]["status"] = "researching"
        logger.info(f"[{job_id}] Fetching market data via FRED/Census/BLS...")

        market_data = run_full_research(
            derived.asset_type, derived.city, derived.state,
        )

        # --- Rent Prediction BEFORE pro forma (Fix #1) ---
        # Predicted growth rates feed INTO the financial model
        rent_prediction = None
        if deal.enable_rent_prediction:
            jobs[job_id]["status"] = "predicting_rents"
            logger.info(f"[{job_id}] Running rent growth prediction...")
            try:
                from models.rent_predictor import RentPredictor
                from services.api_clients.fred_client import FREDClient
                fred = FREDClient()
                cpi_shelter = fred.get_cpi_shelter(limit=60)
                predictor = RentPredictor()
                predictor.train(cpi_shelter.get("annual_growth_rates", []))
                rent_prediction = predictor.predict(
                    deal.hold_period_years,
                    deal.in_place_rent,
                    deal.total_units,
                )
                if rent_prediction and not rent_prediction.get("error"):
                    # Fix #1 & #6: Inject predicted growth rates into deal inputs
                    # so build_pro_forma uses variable year-by-year rates
                    predicted_rates = rent_prediction.get("predicted_rates", [])
                    if predicted_rates:
                        deal.yearly_revenue_growth = predicted_rates
                        logger.info(f"[{job_id}] Injected {len(predicted_rates)} predicted growth rates into pro forma")
                else:
                    logger.warning(f"[{job_id}] Rent prediction warning: {rent_prediction.get('error', 'unknown')}")
            except Exception as e:
                logger.error(f"[{job_id}] Rent prediction failed: {e}")
                rent_prediction = {"error": str(e)}

        # --- Build Pro Forma (now uses variable growth if rent prediction ran) ---
        jobs[job_id]["status"] = "modeling"
        logger.info(f"[{job_id}] Building financial model...")

        pro_forma = build_pro_forma(deal)

        # --- ML Valuation (optional) ---
        ml_valuation = None
        if deal.enable_ml_valuation:
            jobs[job_id]["status"] = "ml_valuation"
            logger.info(f"[{job_id}] Running ML property valuation...")
            try:
                from models.ml_valuation import PropertyValuationModel
                ml_model = PropertyValuationModel()
                ml_model.train(market_data)
                ml_valuation = ml_model.predict(deal, derived, market_data)
                if ml_valuation.get("error"):
                    logger.warning(f"[{job_id}] ML valuation warning: {ml_valuation['error']}")
            except Exception as e:
                logger.error(f"[{job_id}] ML valuation failed: {e}")
                ml_valuation = {"error": str(e)}

        # --- Lease Analysis (Fix #8: multi-file support) ---
        lease_analysis = None
        if deal.lease_pdf_paths:
            jobs[job_id]["status"] = "analyzing_lease"
            logger.info(f"[{job_id}] Analyzing {len(deal.lease_pdf_paths)} lease document(s)...")
            try:
                from services.lease_analyzer import LeaseAnalyzer
                analyzer = LeaseAnalyzer()
                if len(deal.lease_pdf_paths) == 1:
                    lease_analysis = analyzer.analyze_lease(deal.lease_pdf_paths[0])
                else:
                    lease_analysis = analyzer.analyze_multiple_leases(deal.lease_pdf_paths)

                # Fix #1: Compare lease-extracted rent vs user input
                if lease_analysis and not lease_analysis.get("error"):
                    lease_analysis["input_comparison"] = _compare_lease_to_inputs(
                        lease_analysis, deal
                    )

                if lease_analysis and lease_analysis.get("error"):
                    logger.warning(f"[{job_id}] Lease analysis warning: {lease_analysis['error']}")
            except Exception as e:
                logger.error(f"[{job_id}] Lease analysis failed: {e}")
                lease_analysis = {"error": str(e)}

        # --- Sensitivity Tables (Fix #7) ---
        jobs[job_id]["status"] = "sensitivity"
        logger.info(f"[{job_id}] Running sensitivity analysis...")
        sensitivity = _build_sensitivity(deal, pro_forma)

        # --- Integrated Recommendation (Fix #1) ---
        recommendation = _build_recommendation(pro_forma, ml_valuation, lease_analysis, rent_prediction)

        jobs[job_id]["status"] = "generating_excel"
        logger.info(f"[{job_id}] Generating Excel...")

        excel_path = generate_excel(pro_forma, market_data, job_id,
                                    ml_valuation=ml_valuation,
                                    lease_analysis=lease_analysis,
                                    rent_prediction=rent_prediction,
                                    sensitivity=sensitivity)

        jobs[job_id]["status"] = "generating_word"
        logger.info(f"[{job_id}] Generating Word memo...")

        word_path = generate_word(pro_forma, market_data, job_id,
                                  ml_valuation=ml_valuation,
                                  lease_analysis=lease_analysis,
                                  rent_prediction=rent_prediction,
                                  sensitivity=sensitivity)

        jobs[job_id].update({
            "status": "complete",
            "results": pro_forma,
            "market_data": market_data,
            "ml_valuation": ml_valuation,
            "lease_analysis": lease_analysis,
            "rent_prediction": rent_prediction,
            "sensitivity": sensitivity,
            "recommendation": recommendation,
            "excel_path": excel_path,
            "word_path": word_path,
        })
        logger.info(f"[{job_id}] Analysis complete!")

    except Exception as e:
        logger.exception(f"[{job_id}] Analysis failed")
        jobs[job_id].update({
            "status": "error",
            "message": str(e),
        })


def _compare_lease_to_inputs(lease_data: dict, deal: DealInputs) -> dict:
    """Compare lease-extracted rent to user-input rent and flag discrepancies."""
    comparison = {"flags": []}

    # Handle single or multi-lease
    if "individual_leases" in lease_data:
        # Multi-lease: compare portfolio total
        total_monthly = lease_data.get("portfolio_summary", {}).get("total_monthly_rent")
        if total_monthly and deal.in_place_rent > 0 and deal.total_units > 0:
            expected_total = deal.in_place_rent * deal.total_units
            diff_pct = abs(total_monthly - expected_total) / expected_total * 100
            if diff_pct > 10:
                comparison["flags"].append(
                    f"Lease portfolio total rent (${total_monthly:,.0f}/mo) differs from "
                    f"underwriting assumption (${expected_total:,.0f}/mo) by {diff_pct:.0f}%"
                )
            comparison["lease_total_monthly"] = total_monthly
            comparison["input_total_monthly"] = expected_total
    else:
        # Single lease
        lease_rent = lease_data.get("monthly_rent")
        if lease_rent and deal.in_place_rent > 0:
            diff_pct = abs(lease_rent - deal.in_place_rent) / deal.in_place_rent * 100
            if diff_pct > 10:
                comparison["flags"].append(
                    f"Lease rent (${lease_rent:,.0f}/mo) differs from in-place rent "
                    f"assumption (${deal.in_place_rent:,.0f}/mo) by {diff_pct:.0f}%"
                )
            comparison["lease_monthly"] = lease_rent
            comparison["input_monthly"] = deal.in_place_rent

        escalation = lease_data.get("annual_escalation_pct")
        if escalation and deal.revenue_growth_rate > 0:
            growth_pct = deal.revenue_growth_rate * 100
            if abs(escalation - growth_pct) > 1.0:
                comparison["flags"].append(
                    f"Lease escalation ({escalation}%/yr) differs from assumed revenue growth ({growth_pct:.1f}%/yr)"
                )

    return comparison


def _build_sensitivity(deal, pro_forma: dict) -> dict:
    """Build all sensitivity tables."""
    from models.metrics import (
        sensitivity_table_exit_cap, sensitivity_table_interest_rate,
        sensitivity_table_rent_growth, sensitivity_table_purchase_price,
    )

    rev = pro_forma["reversion"]
    derived = pro_forma["inputs"]["derived"]

    return {
        "exit_cap": sensitivity_table_exit_cap(
            base_noi_at_exit=rev["forward_noi"],
            base_exit_cap=rev["exit_cap_rate"],
            equity_invested=derived["equity_required"],
            annual_cfs=pro_forma["annual_btcfs"][:rev["exit_year"]],
            sale_costs_pct=deal.sale_costs_pct,
            loan_balance_at_exit=rev["loan_balance"],
        ),
        "interest_rate": sensitivity_table_interest_rate(deal),
        "rent_growth": sensitivity_table_rent_growth(deal),
        "purchase_price": sensitivity_table_purchase_price(deal),
    }


def _build_recommendation(pro_forma: dict, ml_valuation=None,
                          lease_analysis=None, rent_prediction=None) -> dict:
    """Build integrated recommendation using all available signals (Fix #1)."""
    m = pro_forma["metrics"]
    irr = m.get("levered_irr")
    signals = []
    score = 0

    # 1. IRR signal (primary)
    if irr is not None:
        if irr >= 0.15:
            signals.append(("IRR", "STRONG BUY", f"{irr*100:.1f}% exceeds 15% threshold"))
            score += 2
        elif irr >= 0.12:
            signals.append(("IRR", "BUY", f"{irr*100:.1f}% exceeds 12% threshold"))
            score += 1
        elif irr >= 0.08:
            signals.append(("IRR", "HOLD", f"{irr*100:.1f}% — moderate returns"))
            score += 0
        else:
            signals.append(("IRR", "PASS", f"{irr*100:.1f}% below 8% minimum"))
            score -= 2

    # 2. DSCR signal
    dscr = m.get("dscr_yr1", 0)
    if dscr < 1.0:
        signals.append(("DSCR", "WARNING", f"{dscr:.2f}x — negative cash flow"))
        score -= 2
    elif dscr < 1.25:
        signals.append(("DSCR", "CAUTION", f"{dscr:.2f}x — thin coverage"))
        score -= 1

    # 3. ML valuation signal (Fix #1)
    if ml_valuation and not ml_valuation.get("error"):
        assessment = ml_valuation.get("assessment", "")
        discount = ml_valuation.get("premium_discount_pct", 0)
        if assessment == "UNDERVALUED":
            signals.append(("ML Valuation", "BUY", f"ML model shows {abs(discount):.1f}% discount to predicted value"))
            score += 1
        elif assessment == "OVERVALUED":
            signals.append(("ML Valuation", "CAUTION", f"ML model shows {discount:.1f}% premium to predicted value"))
            score -= 1
        else:
            signals.append(("ML Valuation", "NEUTRAL", "Price is within fair value range"))

    # 4. Lease analysis signal (Fix #1)
    if lease_analysis and not lease_analysis.get("error"):
        risk_flags = lease_analysis.get("risk_flags", [])
        comparison = lease_analysis.get("input_comparison", {})
        comp_flags = comparison.get("flags", [])
        if comp_flags:
            signals.append(("Lease Analysis", "WARNING", "; ".join(comp_flags)))
            score -= 1
        if len(risk_flags) > 3:
            signals.append(("Lease Risk", "CAUTION", f"{len(risk_flags)} risk flags identified"))
            score -= 1

    # 5. Rent prediction signal (Fix #1)
    if rent_prediction and not rent_prediction.get("error"):
        avg_growth = rent_prediction.get("avg_predicted_growth", 0)
        if avg_growth > 4.0:
            signals.append(("Rent Forecast", "POSITIVE", f"{avg_growth:.1f}% avg predicted growth supports returns"))
            score += 1
        elif avg_growth < 1.0:
            signals.append(("Rent Forecast", "CAUTION", f"{avg_growth:.1f}% predicted growth — weak rent outlook"))
            score -= 1

    # Final recommendation
    if score >= 2:
        recommendation = "STRONG BUY"
    elif score >= 1:
        recommendation = "BUY"
    elif score >= 0:
        recommendation = "HOLD / CONDITIONAL"
    else:
        recommendation = "PASS"

    return {
        "recommendation": recommendation,
        "score": score,
        "signals": signals,
        "used_variable_growth": m.get("used_variable_growth", False),
    }


# --- Routes ---

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/analyze", methods=["POST"])
def analyze():
    try:
        deal = _parse_form(request.form)
    except Exception as e:
        return jsonify({"error": f"Invalid input: {e}"}), 400

    if deal.purchase_price <= 0:
        return jsonify({"error": "Purchase price is required"}), 400
    if deal.total_units <= 0:
        return jsonify({"error": "Number of units is required"}), 400
    if deal.in_place_rent <= 0:
        return jsonify({"error": "In-place rent is required"}), 400

    # Handle lease PDF upload (Fix #8: multiple files)
    lease_files = request.files.getlist("lease_pdf")
    saved_paths = []
    for lease_file in lease_files:
        if lease_file and lease_file.filename:
            filename = f"{uuid.uuid4().hex[:8]}_{lease_file.filename}"
            save_path = os.path.join(UPLOAD_DIR, filename)
            lease_file.save(save_path)
            saved_paths.append(save_path)
            logger.info(f"Lease PDF saved to {save_path}")
    if saved_paths:
        deal.lease_pdf_paths = saved_paths

    job_id = uuid.uuid4().hex[:12]
    jobs[job_id] = {"status": "pending", "deal": deal}

    thread = threading.Thread(target=_run_analysis, args=(job_id, deal), daemon=True)
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/processing/<job_id>")
def processing(job_id):
    if job_id not in jobs:
        return render_template("index.html"), 404
    return render_template("processing.html", job_id=job_id)


@app.route("/api/status/<job_id>")
def status(job_id):
    if job_id not in jobs:
        return jsonify({"status": "not_found"}), 404
    job = jobs[job_id]
    resp = {"status": job["status"]}
    if job["status"] == "error":
        resp["message"] = job.get("message", "Unknown error")
    return jsonify(resp)


@app.route("/results/<job_id>")
def results(job_id):
    if job_id not in jobs:
        return render_template("index.html"), 404
    job = jobs[job_id]
    if job["status"] != "complete":
        return render_template("processing.html", job_id=job_id)
    return render_template("results.html",
                           results=job["results"],
                           job_id=job_id,
                           ml_valuation=job.get("ml_valuation"),
                           lease_analysis=job.get("lease_analysis"),
                           rent_prediction=job.get("rent_prediction"),
                           sensitivity=job.get("sensitivity"),
                           recommendation=job.get("recommendation"))


@app.route("/api/results/<job_id>")
def results_json(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Not found"}), 404
    job = jobs[job_id]
    if job["status"] != "complete":
        return jsonify({"status": job["status"]}), 202
    return jsonify(job["results"])


@app.route("/api/download/<job_id>/<file_type>")
def download(job_id, file_type):
    if job_id not in jobs:
        return jsonify({"error": "Not found"}), 404
    job = jobs[job_id]
    if job["status"] != "complete":
        return jsonify({"error": "Analysis not complete"}), 400

    if file_type == "excel":
        path = job.get("excel_path")
        if path and os.path.exists(path):
            return send_file(path, as_attachment=True, download_name=f"underwriting_{job_id}.xlsx")
    elif file_type == "word":
        path = job.get("word_path")
        if path and os.path.exists(path):
            return send_file(path, as_attachment=True, download_name=f"investment_memo_{job_id}.docx")

    return jsonify({"error": "File not found"}), 404


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=DEBUG, use_reloader=False)
