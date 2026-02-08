"""Flask application for Real Estate Investment Underwriting Tool."""

import os
import uuid
import threading
import logging
from flask import Flask, render_template, request, jsonify, send_file

from config import OUTPUT_DIR, PORT, DEBUG, SECRET_KEY
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
    )


def _run_analysis(job_id: str, deal: DealInputs):
    """Run the full analysis pipeline in a background thread."""
    try:
        derived = derive_assumptions(deal)

        jobs[job_id]["status"] = "researching"
        logger.info(f"[{job_id}] Market research for {derived.city}, {derived.state}...")

        market_data = run_full_research(
            derived.asset_type, derived.city, derived.state,
        )

        jobs[job_id]["status"] = "modeling"
        logger.info(f"[{job_id}] Building financial model...")

        pro_forma = build_pro_forma(deal)

        jobs[job_id]["status"] = "generating_excel"
        logger.info(f"[{job_id}] Generating Excel...")

        excel_path = generate_excel(pro_forma, market_data, job_id)

        jobs[job_id]["status"] = "generating_word"
        logger.info(f"[{job_id}] Generating Word memo...")

        word_path = generate_word(pro_forma, market_data, job_id)

        jobs[job_id].update({
            "status": "complete",
            "results": pro_forma,
            "market_data": market_data,
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
    return render_template("results.html", results=job["results"], job_id=job_id)


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
    app.run(host="0.0.0.0", port=PORT, debug=DEBUG)
