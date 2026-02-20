"""Flask application for Real Estate Investment Underwriting Tool."""

import os
import uuid
import threading
import logging
from flask import Flask, render_template, request, jsonify, send_file

from config import OUTPUT_DIR, UPLOAD_DIR, PORT, DEBUG, SECRET_KEY, ANTHROPIC_API_KEY, OPENAI_API_KEY
from models.assumptions import DealInputs, derive_assumptions
from models.unit_mix import parse_unit_mix_from_form
from models.financial_model import build_pro_forma
from services.market_research import run_full_research
from services.excel_generator import generate_excel
from services.word_generator import generate_word
from services.pdf_generator import generate_pdf
from services.ai_memo import generate_ai_memo
from services.database import init_db, save_deal as db_save_deal, load_all_deals, load_deal

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
app.secret_key = SECRET_KEY

# In-memory stores
jobs: dict[str, dict] = {}
saved_deals: dict[str, dict] = {}  # job_id -> saved deal summary for comparison

# Initialize SQLite database
try:
    init_db()
except Exception as _db_err:
    logger.warning(f"DB init failed (non-fatal): {_db_err}")


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
    deal = DealInputs(
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
        # Waterfall
        enable_waterfall=form.get("enable_waterfall") == "on",
        lp_equity_pct=_f(form.get("lp_equity_pct"), 90) / 100,
        preferred_return=_f(form.get("preferred_return"), 8) / 100,
        promote_pct=_f(form.get("promote_pct"), 20) / 100,
    )

    # Parse unit mix; if provided, override blended in_place_rent / market_rent
    unit_mix = parse_unit_mix_from_form(form)
    if not unit_mix.is_empty():
        deal.unit_mix = unit_mix
        # Reconcile total_units with mix if user left total_units blank
        if deal.total_units == 0:
            deal.total_units = unit_mix.total_units
        # Override average rents with blended values from the detailed mix
        if unit_mix.blended_in_place_rent > 0:
            deal.in_place_rent = round(unit_mix.blended_in_place_rent, 2)
        if unit_mix.blended_market_rent > 0:
            deal.market_rent = round(unit_mix.blended_market_rent, 2)

    return deal


def _run_analysis(job_id: str, deal: DealInputs):
    """Run the full analysis pipeline in a background thread."""
    try:
        derived = derive_assumptions(deal)

        jobs[job_id]["status"] = "researching"
        logger.info(f"[{job_id}] Fetching market data via FRED/Census/BLS...")

        # Estimate avg bedrooms and sq_ft for RentCast AVM
        avg_sq_ft = int(deal.total_sf / deal.total_units) if deal.total_units > 0 and deal.total_sf > 0 else None

        market_data = run_full_research(
            derived.asset_type, derived.city, derived.state,
            address=deal.address,
            zip_code=derived.zip_code,
            avg_bedrooms=2,
            avg_sq_ft=avg_sq_ft,
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
                # Get ZORI growth rates for the city if available
                zori_rates = market_data.get("rent_trends", {}).get("zori_growth_rates", [])
                predictor = RentPredictor()
                predictor.train(cpi_shelter.get("annual_growth_rates", []),
                                zori_growth_rates=zori_rates)
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

        # --- Backtest (after rent prediction, before pro forma) ---
        backtest_result = None
        if rent_prediction and not rent_prediction.get("error"):
            jobs[job_id]["status"] = "backtesting"
            logger.info(f"[{job_id}] Running rent prediction backtest...")
            try:
                from models.backtest import RentBacktester
                backtester = RentBacktester()
                zori_rates = market_data.get("rent_trends", {}).get("zori_growth_rates", [])
                backtest_result = backtester.run_backtest(
                    cpi_shelter.get("annual_growth_rates", []),
                    zori_growth_rates=zori_rates,
                )
                if backtest_result.get("error"):
                    logger.warning(f"[{job_id}] Backtest warning: {backtest_result['error']}")
            except Exception as e:
                logger.error(f"[{job_id}] Backtest failed: {e}")
                backtest_result = {"error": str(e)}

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

        # --- Monte Carlo Simulation ---
        monte_carlo = None
        jobs[job_id]["status"] = "monte_carlo"
        logger.info(f"[{job_id}] Running Monte Carlo simulation...")
        try:
            from models.monte_carlo import MonteCarloSimulator
            mc_sim = MonteCarloSimulator(n_iterations=1000)
            monte_carlo = mc_sim.run(deal, build_pro_forma)
            if monte_carlo.get("error"):
                logger.warning(f"[{job_id}] Monte Carlo warning: {monte_carlo['error']}")
        except Exception as e:
            logger.error(f"[{job_id}] Monte Carlo failed: {e}")
            monte_carlo = {"error": str(e)}

        # --- Integrated Recommendation (Fix #1) ---
        recommendation = _build_recommendation(pro_forma, ml_valuation, lease_analysis, rent_prediction, monte_carlo)

        # --- Waterfall / LP-GP Promote ---
        waterfall = None
        if deal.enable_waterfall:
            jobs[job_id]["status"] = "waterfall"
            logger.info(f"[{job_id}] Running LP/GP waterfall...")
            try:
                from models.waterfall import run_waterfall
                derived_wf = pro_forma["inputs"]["derived"]
                rev_wf = pro_forma["reversion"]
                waterfall = run_waterfall(
                    lp_equity_pct=deal.lp_equity_pct,
                    preferred_return=deal.preferred_return,
                    promote_pct=deal.promote_pct,
                    equity_invested=derived_wf["equity_required"],
                    annual_btcfs=pro_forma.get("annual_btcfs", [])[:deal.hold_period_years],
                    exit_proceeds=rev_wf.get("net_sale_proceeds", 0),
                    hold_years=deal.hold_period_years,
                )
                if waterfall.get("error"):
                    logger.warning(f"[{job_id}] Waterfall warning: {waterfall['error']}")
            except Exception as e:
                logger.error(f"[{job_id}] Waterfall failed: {e}")
                waterfall = {"error": str(e)}

        # --- AI Deal Memo (Claude-written narrative grounded in deal data) ---
        ai_memo = None
        if OPENAI_API_KEY:
            jobs[job_id]["status"] = "generating_ai_memo"
            logger.info(f"[{job_id}] Generating AI investment memo...")
            try:
                memo_job = {
                    "deal": deal,
                    "results": pro_forma,
                    "market_data": market_data,
                    "ml_valuation": ml_valuation,
                    "monte_carlo": monte_carlo,
                    "recommendation": recommendation,
                    "rent_prediction": rent_prediction,
                }
                ai_memo = generate_ai_memo(memo_job, OPENAI_API_KEY)
                if ai_memo.get("error"):
                    logger.warning(f"[{job_id}] AI memo warning: {ai_memo['error']}")
            except Exception as e:
                logger.error(f"[{job_id}] AI memo failed: {e}")
                ai_memo = {"error": str(e)}

        # --- Scenario Comparison (Base / Bull / Bear) ---
        jobs[job_id]["status"] = "scenarios"
        logger.info(f"[{job_id}] Running Base/Bull/Bear scenarios...")
        try:
            from models.scenarios import run_scenarios
            scenarios = run_scenarios(deal, build_pro_forma)
        except Exception as e:
            logger.error(f"[{job_id}] Scenario comparison failed: {e}")
            scenarios = {"error": str(e)}

        # Serialize unit mix (needs market_data for HUD FMR comparison)
        unit_mix_data = None
        if deal.unit_mix and not deal.unit_mix.is_empty():
            hud_fmr = market_data.get("hud_fmr", {}) if market_data else {}
            unit_mix_data = deal.unit_mix.with_hud_fmr(hud_fmr)

        jobs[job_id]["status"] = "generating_excel"
        logger.info(f"[{job_id}] Generating Excel...")

        excel_path = generate_excel(pro_forma, market_data, job_id,
                                    ml_valuation=ml_valuation,
                                    lease_analysis=lease_analysis,
                                    rent_prediction=rent_prediction,
                                    sensitivity=sensitivity,
                                    backtest=backtest_result,
                                    monte_carlo=monte_carlo)

        jobs[job_id]["status"] = "generating_word"
        logger.info(f"[{job_id}] Generating Word memo...")

        word_path = generate_word(pro_forma, market_data, job_id,
                                  ml_valuation=ml_valuation,
                                  lease_analysis=lease_analysis,
                                  rent_prediction=rent_prediction,
                                  sensitivity=sensitivity,
                                  backtest=backtest_result,
                                  monte_carlo=monte_carlo,
                                  ai_memo=ai_memo,
                                  unit_mix=unit_mix_data)

        jobs[job_id]["status"] = "generating_pdf"
        logger.info(f"[{job_id}] Generating PDF report...")

        pdf_path = generate_pdf(pro_forma, market_data, job_id,
                                ml_valuation=ml_valuation,
                                lease_analysis=lease_analysis,
                                rent_prediction=rent_prediction,
                                sensitivity=sensitivity,
                                backtest=backtest_result,
                                monte_carlo=monte_carlo)

        jobs[job_id].update({
            "status": "complete",
            "results": pro_forma,
            "market_data": market_data,
            "ml_valuation": ml_valuation,
            "lease_analysis": lease_analysis,
            "rent_prediction": rent_prediction,
            "backtest": backtest_result,
            "monte_carlo": monte_carlo,
            "sensitivity": sensitivity,
            "recommendation": recommendation,
            "ai_memo": ai_memo,
            "unit_mix": unit_mix_data,
            "waterfall": waterfall,
            "scenarios": scenarios,
            "excel_path": excel_path,
            "word_path": word_path,
            "pdf_path": pdf_path,
            "deal": deal,
        })
        logger.info(f"[{job_id}] Analysis complete!")

        # Persist to SQLite (non-blocking, non-fatal)
        try:
            db_save_deal(job_id, jobs[job_id])
        except Exception as _e:
            logger.warning(f"[{job_id}] DB save failed (non-fatal): {_e}")

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
                          lease_analysis=None, rent_prediction=None,
                          monte_carlo=None) -> dict:
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

    # 6. Monte Carlo signal
    if monte_carlo and not monte_carlo.get("error"):
        mc_signal = monte_carlo.get("mc_signal", "NEUTRAL")
        mc_detail = monte_carlo.get("mc_detail", "")
        if mc_signal == "POSITIVE":
            signals.append(("Monte Carlo", "POSITIVE", mc_detail))
            score += 1
        elif mc_signal == "WARNING":
            signals.append(("Monte Carlo", "WARNING", mc_detail))
            score -= 1
        else:
            signals.append(("Monte Carlo", "NEUTRAL", mc_detail))

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
                           market_data=job.get("market_data"),
                           ml_valuation=job.get("ml_valuation"),
                           lease_analysis=job.get("lease_analysis"),
                           rent_prediction=job.get("rent_prediction"),
                           backtest=job.get("backtest"),
                           monte_carlo=job.get("monte_carlo"),
                           sensitivity=job.get("sensitivity"),
                           recommendation=job.get("recommendation"),
                           ai_memo=job.get("ai_memo"),
                           unit_mix=job.get("unit_mix"),
                           waterfall=job.get("waterfall"),
                           scenarios=job.get("scenarios"))


@app.route("/api/results/<job_id>")
def results_json(job_id):
    if job_id not in jobs:
        return jsonify({"error": "Not found"}), 404
    job = jobs[job_id]
    if job["status"] != "complete":
        return jsonify({"status": job["status"]}), 202
    return jsonify(job["results"])


@app.route("/api/download/<job_id>/file/<file_type>")
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
    elif file_type == "pdf":
        path = job.get("pdf_path")
        if path and os.path.exists(path):
            return send_file(path, as_attachment=True, download_name=f"investment_report_{job_id}.pdf")

    return jsonify({"error": "File not found"}), 404


@app.route("/api/save_deal/<job_id>", methods=["POST"])
def save_deal(job_id):
    """Save a completed deal for comparison."""
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    if job["status"] != "complete":
        return jsonify({"error": "Analysis not complete"}), 400

    results = job["results"]
    m = results["metrics"]
    deal = results["inputs"]["deal"]
    derived = results["inputs"]["derived"]

    saved_deals[job_id] = {
        "job_id": job_id,
        "property_name": derived["property_name"],
        "address": deal["address"],
        "property_type": deal["property_type"],
        "purchase_price": deal["purchase_price"],
        "total_units": deal["total_units"],
        "price_per_unit": derived["price_per_unit"],
        "current_noi": deal["current_noi"],
        "in_place_rent": deal["in_place_rent"],
        "market_rent": deal["market_rent"],
        "occupancy": deal["occupancy"],
        "levered_irr": m.get("levered_irr"),
        "after_tax_irr": m.get("after_tax_irr"),
        "equity_multiple": m.get("equity_multiple"),
        "cash_on_cash_yr1": m.get("cash_on_cash_yr1"),
        "dscr_yr1": m.get("dscr_yr1"),
        "going_in_cap_rate": m.get("going_in_cap_rate"),
        "exit_cap_rate": m.get("exit_cap_rate"),
        "yield_on_cost": m.get("yield_on_cost"),
        "stabilized_yoc": m.get("stabilized_yoc"),
        "annual_nois": results.get("annual_nois", []),
        "ml_assessment": job.get("ml_valuation", {}).get("assessment") if job.get("ml_valuation") and not job.get("ml_valuation", {}).get("error") else None,
        "mc_prob_12": job.get("monte_carlo", {}).get("probabilities", {}).get("IRR > 12%") if job.get("monte_carlo") and not job.get("monte_carlo", {}).get("error") else None,
    }
    return jsonify({"saved": True, "deal_count": len(saved_deals)})


@app.route("/api/saved_deals")
def get_saved_deals():
    """Get all saved deals for comparison."""
    return jsonify(list(saved_deals.values()))


@app.route("/compare")
def compare():
    """Show comparison page."""
    return render_template("compare.html", deals=list(saved_deals.values()))


@app.route("/api/whatif/<job_id>", methods=["POST"])
def whatif(job_id):
    """Interactive what-if analysis endpoint."""
    import copy

    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    if job["status"] != "complete":
        return jsonify({"error": "Analysis not complete"}), 400

    deal_obj = job.get("deal")
    if deal_obj is None:
        return jsonify({"error": "Deal data not available"}), 400

    data = request.get_json()
    if not data:
        return jsonify({"error": "No parameters provided"}), 400

    sim_deal = copy.deepcopy(deal_obj)
    sim_deal.yearly_revenue_growth = []  # Use flat rate for what-if

    if "rent_growth" in data:
        sim_deal.revenue_growth_rate = data["rent_growth"] / 100
    if "exit_cap_spread" in data:
        sim_deal.exit_cap_rate_spread = data["exit_cap_spread"] / 10000
    if "occupancy" in data:
        sim_deal.occupancy = data["occupancy"] / 100
    if "expense_growth" in data:
        sim_deal.expense_growth_rate = data["expense_growth"] / 100

    try:
        result = build_pro_forma(sim_deal)
        m = result["metrics"]
        return jsonify({
            "levered_irr": round(m["levered_irr"] * 100, 1) if m.get("levered_irr") else None,
            "equity_multiple": round(m["equity_multiple"], 2) if m.get("equity_multiple") else None,
            "dscr_yr1": round(m["dscr_yr1"], 2) if m.get("dscr_yr1") else None,
            "cash_on_cash_yr1": round(m["cash_on_cash_yr1"] * 100, 1) if m.get("cash_on_cash_yr1") else None,
            "yield_on_cost": round(m["yield_on_cost"] * 100, 2) if m.get("yield_on_cost") else None,
            "annual_nois": result.get("annual_nois", []),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def _build_chat_system_prompt(job: dict) -> str:
    """Build a grounded system prompt for the deal chat interface.

    Extracts all computed results from the job and structures them as
    factual context for Claude. Claude must answer only from this data.
    """
    deal = job.get("deal")
    pf = job.get("results", {})
    metrics = pf.get("metrics", {})
    reversion = pf.get("reversion", {})
    pro_forma_rows = pf.get("pro_forma", [])
    market = job.get("market_data", {})
    ml = job.get("ml_valuation", {})
    mc = job.get("monte_carlo", {})
    rec = job.get("recommendation", {})
    lease = job.get("lease_analysis", {})

    def pct(v):
        return f"{v*100:.1f}%" if v is not None else "N/A"

    def dollar(v):
        return f"${v:,.0f}" if v is not None else "N/A"

    def fmt(v, decimals=2):
        return f"{v:.{decimals}f}" if v is not None else "N/A"

    # --- Deal basics ---
    lines = [
        "You are an expert real estate investment analyst and deal assistant.",
        "You have full access to the underwriting results for the following deal.",
        "Answer all questions using ONLY the data provided below — never fabricate numbers.",
        "Be concise, precise, and quantitative. Flag uncertainty where it exists.",
        "",
        "---------------------------------------",
        "DEAL OVERVIEW",
        "---------------------------------------",
    ]

    if deal:
        lines += [
            f"Property Type: {deal.property_type}",
            f"Address: {deal.address}",
            f"Purchase Price: {dollar(deal.purchase_price)}",
            f"Total Units: {deal.total_units}",
            f"In-Place Rent: {dollar(deal.in_place_rent)}/unit/month",
            f"Occupancy: {pct(deal.occupancy)}",
            f"Hold Period: {deal.hold_period_years} years",
            f"LTV: {pct(deal.ltv)}",
            f"Interest Rate: {pct(deal.interest_rate)}",
            f"Revenue Growth Rate: {pct(deal.revenue_growth_rate)}",
        ]

    # --- Key metrics ---
    lines += [
        "",
        "---------------------------------------",
        "KEY FINANCIAL METRICS",
        "---------------------------------------",
        f"Levered IRR: {pct(metrics.get('levered_irr'))}",
        f"Unlevered IRR: {pct(metrics.get('unlevered_irr'))}",
        f"After-Tax IRR: {pct(metrics.get('after_tax_irr'))}",
        f"Equity Multiple: {fmt(metrics.get('equity_multiple'))}x",
        f"Cash-on-Cash (Yr 1): {pct(metrics.get('cash_on_cash_yr1'))}",
        f"DSCR (Yr 1): {fmt(metrics.get('dscr_yr1'))}",
        f"Yield on Cost: {pct(metrics.get('yield_on_cost'))}",
        f"Going-In Cap Rate: {pct(metrics.get('going_in_cap_rate'))}",
        f"Exit Cap Rate: {pct(metrics.get('exit_cap_rate'))}",
        f"Price per Unit: {dollar(metrics.get('price_per_unit'))}",
    ]

    # --- Exit ---
    if reversion:
        lines += [
            "",
            "---------------------------------------",
            "EXIT / REVERSION",
            "---------------------------------------",
            f"Exit Year: {reversion.get('exit_year', 'N/A')}",
            f"Forward NOI at Exit: {dollar(reversion.get('forward_noi'))}",
            f"Exit Sale Price: {dollar(reversion.get('sale_price'))}",
            f"Net Sale Proceeds (after tax): {dollar(reversion.get('net_sale_proceeds_after_tax'))}",
        ]

    # --- Pro forma summary (first 3 years) ---
    if pro_forma_rows:
        lines += [
            "",
            "---------------------------------------",
            "PRO FORMA SUMMARY (first 3 years)",
            "---------------------------------------",
        ]
        for row in pro_forma_rows[:3]:
            yr = row.get("year", "?")
            egi = row.get("effective_gross_income", 0)
            exp = row.get("total_expenses", 0)
            noi = row.get("noi", 0)
            btcf = row.get("btcf", 0)
            lines.append(
                f"  Year {yr}: EGI={dollar(egi)}  Expenses={dollar(exp)}"
                f"  NOI={dollar(noi)}  BTCF={dollar(btcf)}"
            )

    # --- Market data ---
    demo = market.get("demographics", {}).get("structured", {}) if market else {}
    caps = market.get("cap_rates", {}) if market else {}
    crime = market.get("crime", {}) if market else {}
    ws = market.get("walkscore", {}) if market else {}
    hud = market.get("hud_fmr", {}) if market else {}
    rc = market.get("rentcast", {}) if market else {}
    rc_stats = rc.get("market_stats", {}) if rc else {}
    rc_avm = rc.get("rent_estimate", {}) if rc else {}

    lines += [
        "",
        "---------------------------------------",
        "MARKET DATA",
        "---------------------------------------",
    ]

    if demo:
        lines += [
            f"Population: {demo.get('population', 'N/A'):,}" if demo.get('population') else "Population: N/A",
            f"Median Household Income: {dollar(demo.get('median_income'))}",
            f"Median Gross Rent: {dollar(demo.get('median_rent'))}/mo",
            f"Unemployment Rate: {fmt(demo.get('unemployment_rate'))}%",
            f"Renter-Occupied: {demo.get('renter_pct', 'N/A')}%",
        ]

    if caps.get("average_cap_rate"):
        lines.append(f"Market Cap Rate (FRED-derived): {fmt(caps['average_cap_rate'])}%")
    if caps.get("treasury_10yr"):
        lines.append(f"10-Year Treasury: {fmt(caps['treasury_10yr'])}%")
    if caps.get("mortgage_30yr"):
        lines.append(f"30-Year Mortgage Rate: {fmt(caps['mortgage_30yr'])}%")

    if crime.get("available"):
        lines += [
            f"Crime Risk: {crime.get('risk_level')} — {crime.get('violent_per_100k')} violent crimes/100k",
            f"vs. National Average: {crime.get('vs_national_pct', 0):+.1f}%",
        ]

    if ws.get("available"):
        lines += [
            f"Walk Score: {ws.get('walk_score')} ({ws.get('walk_label')})",
            f"Transit Score: {ws.get('transit_score')} ({ws.get('transit_label')})",
            f"Bike Score: {ws.get('bike_score')} ({ws.get('bike_label')})",
        ]

    fz = market.get("flood_zone", {}) if market else {}
    if fz.get("available"):
        lines += [
            f"FEMA Flood Zone: {fz.get('flood_zone')} — {fz.get('risk_level')}",
            f"SFHA (insurance required): {'Yes' if fz.get('sfha') else 'No'}",
            f"Description: {fz.get('description', '')}",
        ]

    if hud.get("available"):
        fmr_br = hud.get("fmr_by_bedroom", {})
        lines += [
            f"HUD Fair Market Rent (2026): Studio=${fmr_br.get('efficiency', 'N/A')}"
            f" | 1BR=${fmr_br.get('one_br', 'N/A')}"
            f" | 2BR=${fmr_br.get('two_br', 'N/A')}"
            f" | 3BR=${fmr_br.get('three_br', 'N/A')}",
            f"HUD Metro: {hud.get('metro_name')}",
        ]
        if hud.get("matched_fmr") and deal and deal.in_place_rent:
            diff = (deal.in_place_rent - hud["matched_fmr"]) / hud["matched_fmr"] * 100
            lines.append(f"In-Place Rent vs HUD FMR: {diff:+.1f}%")

    if rc_stats.get("available"):
        lines += [
            f"RentCast Market (Zip): Avg Rent=${rc_stats.get('average_rent', 'N/A')}/mo"
            f"  Median=${rc_stats.get('median_rent', 'N/A')}/mo"
            f"  Days on Market={rc_stats.get('avg_days_on_market', 'N/A')}",
        ]
    if rc_avm.get("available") and rc_avm.get("estimated_rent"):
        lines.append(f"RentCast AVM Estimate: ${rc_avm['estimated_rent']:,.0f}/mo"
                     f" (range: ${rc_avm.get('rent_range_low', '?'):,}"
                     f"–${rc_avm.get('rent_range_high', '?'):,})")

    # --- ML valuation ---
    if ml and ml.get("available"):
        lines += [
            "",
            "---------------------------------------",
            "ML VALUATION (GradientBoosting)",
            "---------------------------------------",
            f"Assessment: {ml.get('assessment')}",
            f"Predicted Value: {dollar(ml.get('predicted_value'))}",
            f"Actual Price: {dollar(ml.get('actual_price'))}",
            f"Premium/Discount: {ml.get('premium_discount_pct', 0):+.1f}%",
            f"Model R²: {fmt(ml.get('test_r2'))}",
        ]

    # --- Monte Carlo ---
    if mc and mc.get("irr_mean") is not None:
        probs = mc.get("probabilities", {})
        lines += [
            "",
            "---------------------------------------",
            "MONTE CARLO (1,000 simulations)",
            "---------------------------------------",
            f"IRR: P10={pct(mc.get('irr_p10'))}  P50={pct(mc.get('irr_p50'))}  "
            f"P90={pct(mc.get('irr_p90'))}",
            f"Mean IRR: {pct(mc.get('irr_mean'))}  Std Dev: {pct(mc.get('irr_std'))}",
        ]
        for label, val in probs.items():
            lines.append(f"  {label}: {val:.1f}%")

    # --- Recommendation ---
    if rec:
        lines += [
            "",
            "---------------------------------------",
            "INTEGRATED RECOMMENDATION",
            "---------------------------------------",
            f"Signal: {rec.get('recommendation', 'N/A')}",
            f"Score: {rec.get('score', 'N/A')}",
        ]
        for sig in rec.get("signals", []):
            lines.append(f"  • {sig}")

    # --- Lease ---
    if lease and not lease.get("error") and lease.get("portfolio_summary"):
        ps = lease["portfolio_summary"]
        lines += [
            "",
            "---------------------------------------",
            "LEASE ANALYSIS",
            "---------------------------------------",
            f"Leases Analyzed: {lease.get('lease_count', 'N/A')}",
            f"Weighted Avg Escalation: {ps.get('weighted_avg_escalation', 'N/A')}%/yr",
            f"Risk Flags: {', '.join(ps.get('risk_flags', [])) or 'None identified'}",
        ]

    lines += [
        "",
        "---------------------------------------",
        "INSTRUCTIONS",
        "---------------------------------------",
        "- Answer only from the data above. Do not invent numbers.",
        "- If asked about something not in the data, say so clearly.",
        "- Be concise: 2-4 sentences for most answers.",
        "- For numerical questions, cite the exact figure from the data.",
        "- For 'what if' questions, reason from the financial relationships",
        "  shown (e.g., lower occupancy → lower NOI → lower IRR) but note",
        "  you cannot recalculate without running the full model.",
    ]

    return "\n".join(lines)


@app.route("/api/chat/<job_id>", methods=["POST"])
def chat(job_id):
    """AI deal chat endpoint — Claude grounded in the deal's analysis results."""
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    if job["status"] != "complete":
        return jsonify({"error": "Analysis not complete yet"}), 400

    data = request.get_json()
    if not data or not data.get("message"):
        return jsonify({"error": "No message provided"}), 400

    history = data.get("history", [])   # list of {role, content}
    user_message = data["message"].strip()

    if not OPENAI_API_KEY:
        return jsonify({"error": "OPENAI_API_KEY not configured"}), 500

    try:
        from openai import OpenAI
        client = OpenAI(api_key=OPENAI_API_KEY)

        system_prompt = _build_chat_system_prompt(job)

        messages = [{"role": "system", "content": system_prompt}]
        for h in history[-10:]:
            if h.get("role") in ("user", "assistant") and h.get("content"):
                messages.append({"role": h["role"], "content": h["content"]})
        messages.append({"role": "user", "content": user_message})

        response = client.chat.completions.create(
            model="gpt-4o",
            max_tokens=1024,
            messages=messages,
        )
        reply = response.choices[0].message.content
        return jsonify({"response": reply})

    except Exception as e:
        logger.error(f"Chat error for job {job_id}: {e}")
        return jsonify({"error": f"AI error: {str(e)}"}), 500


@app.route("/api/past_deals")
def past_deals():
    """Return all persisted deals from SQLite."""
    return jsonify(load_all_deals())


@app.route("/api/reload_deal/<job_id>", methods=["POST"])
def reload_deal(job_id):
    """Reload a persisted deal from SQLite into the in-memory jobs dict."""
    if job_id in jobs and jobs[job_id].get("status") == "complete":
        return jsonify({"loaded": True, "source": "memory"})
    stored = load_deal(job_id)
    if stored:
        stored["status"] = "complete"
        jobs[job_id] = stored
        return jsonify({"loaded": True, "source": "database"})
    return jsonify({"loaded": False, "error": "Not found in database"}), 404


@app.route("/api/generate_memo/<job_id>", methods=["POST"])
def regenerate_memo(job_id):
    """Regenerate the AI investment memo on demand."""
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    if job["status"] != "complete":
        return jsonify({"error": "Analysis not complete yet"}), 400

    if not OPENAI_API_KEY:
        return jsonify({"error": "OPENAI_API_KEY not configured"}), 500

    memo_job = {
        "deal": job.get("deal"),
        "results": job.get("results"),
        "market_data": job.get("market_data"),
        "ml_valuation": job.get("ml_valuation"),
        "monte_carlo": job.get("monte_carlo"),
        "recommendation": job.get("recommendation"),
        "rent_prediction": job.get("rent_prediction"),
    }

    try:
        ai_memo = generate_ai_memo(memo_job, OPENAI_API_KEY)
        jobs[job_id]["ai_memo"] = ai_memo
        return jsonify(ai_memo)
    except Exception as e:
        logger.error(f"Memo regeneration error for job {job_id}: {e}")
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, debug=DEBUG, use_reloader=False)
