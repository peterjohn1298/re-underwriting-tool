"""Generate professional PDF investment report using fpdf2."""

import os
import logging
from datetime import datetime
from fpdf import FPDF

from config import OUTPUT_DIR

logger = logging.getLogger(__name__)

# Colors
NAVY = (27, 42, 74)
DARK_BLUE = (44, 62, 107)
GOLD = (200, 169, 81)
WHITE = (255, 255, 255)
LIGHT_GRAY = (242, 242, 242)
BLACK = (0, 0, 0)


class REUnderwritingPDF(FPDF):
    """Custom PDF with header/footer for RE reports."""

    def __init__(self, property_name=""):
        super().__init__()
        self.property_name = property_name

    def header(self):
        if self.page_no() > 1:
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(*NAVY)
            self.cell(0, 6, f"{self.property_name} | Investment Memorandum", align="R")
            self.ln(8)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Confidential | Page {self.page_no()}", align="C")


def _section(pdf, title):
    """Add a section heading."""
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(*NAVY)
    pdf.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")
    pdf.set_draw_color(*GOLD)
    pdf.line(pdf.l_margin, pdf.get_y(), pdf.w - pdf.r_margin, pdf.get_y())
    pdf.ln(4)


def _subsection(pdf, title):
    """Add a subsection heading."""
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(*DARK_BLUE)
    pdf.cell(0, 8, title, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)


def _kv_table(pdf, items, col1_w=70, col2_w=110):
    """Add a key-value table."""
    pdf.set_font("Helvetica", "", 9)
    for i, (key, val) in enumerate(items):
        if i % 2 == 0:
            pdf.set_fill_color(*LIGHT_GRAY)
        else:
            pdf.set_fill_color(*WHITE)
        pdf.set_text_color(*BLACK)
        pdf.cell(col1_w, 7, str(key), border=0, fill=True)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(col2_w, 7, str(val), border=0, fill=True, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9)
    pdf.ln(3)


def _data_table(pdf, headers, rows, col_widths=None):
    """Add a data table with headers."""
    if not col_widths:
        avail = pdf.w - pdf.l_margin - pdf.r_margin
        col_widths = [avail / len(headers)] * len(headers)

    # Header
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(*WHITE)
    for i, h in enumerate(headers):
        pdf.cell(col_widths[i], 7, h, border=0, fill=True, align="C")
    pdf.ln()

    # Rows
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(*BLACK)
    for ri, row in enumerate(rows):
        if ri % 2 == 0:
            pdf.set_fill_color(*LIGHT_GRAY)
        else:
            pdf.set_fill_color(*WHITE)
        for ci, val in enumerate(row):
            align = "L" if ci == 0 else "R"
            pdf.cell(col_widths[ci], 6, str(val), border=0, fill=True, align=align)
        pdf.ln()
    pdf.ln(3)


def _cur(val):
    if val is None:
        return "N/A"
    if abs(val) >= 1e6:
        return f"${val/1e6:,.1f}M"
    if abs(val) >= 1000:
        return f"${val:,.0f}"
    return f"${val:,.2f}"


def _pct(val):
    return f"{val*100:.2f}%" if val is not None else "N/A"


def generate_pdf(pro_forma: dict, market_data: dict, job_id: str,
                 ml_valuation: dict = None,
                 lease_analysis: dict = None,
                 rent_prediction: dict = None,
                 sensitivity: dict = None,
                 backtest: dict = None,
                 monte_carlo: dict = None) -> str:
    """Generate a PDF investment report."""

    inp = pro_forma["inputs"]
    deal = inp["deal"]
    derived = inp["derived"]
    m = pro_forma["metrics"]
    pf = pro_forma["pro_forma"]
    su = pro_forma["sources_uses"]
    rev = pro_forma["reversion"]

    has_ml = ml_valuation and not ml_valuation.get("error")
    has_lease = lease_analysis and not lease_analysis.get("error")
    has_rent = rent_prediction and not rent_prediction.get("error")
    has_bt = backtest and not backtest.get("error") and has_rent
    has_mc = monte_carlo and not monte_carlo.get("error")

    pdf = REUnderwritingPDF(derived["property_name"])
    pdf.set_auto_page_break(auto=True, margin=20)

    # ===== COVER PAGE =====
    pdf.add_page()
    pdf.ln(60)
    pdf.set_font("Helvetica", "B", 28)
    pdf.set_text_color(*NAVY)
    pdf.cell(0, 15, "INVESTMENT MEMORANDUM", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)
    pdf.set_font("Helvetica", "B", 22)
    pdf.set_text_color(*GOLD)
    pdf.cell(0, 12, derived["property_name"], align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    pdf.set_font("Helvetica", "", 14)
    pdf.set_text_color(*DARK_BLUE)
    pdf.cell(0, 10, deal["address"], align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 10, deal["property_type"], align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(20)
    pdf.set_font("Helvetica", "", 12)
    pdf.set_text_color(*NAVY)
    pdf.cell(0, 10, datetime.now().strftime("%B %d, %Y"), align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(204, 0, 0)
    pdf.cell(0, 10, "CONFIDENTIAL", align="C", new_x="LMARGIN", new_y="NEXT")

    # ===== EXECUTIVE SUMMARY =====
    pdf.add_page()
    _section(pdf, "1. Executive Summary")

    irr_val = m.get("levered_irr")
    rec = "BUY" if irr_val and irr_val >= 0.12 else ("HOLD" if irr_val and irr_val >= 0.08 else "PASS")

    _kv_table(pdf, [
        ("Property", derived["property_name"]),
        ("Type", deal["property_type"]),
        ("Address", deal["address"]),
        ("Year Built", str(deal["year_built"])),
        ("Units / SF", f"{deal['total_units']} units / {deal['total_sf']:,.0f} SF"),
        ("Purchase Price", _cur(deal["purchase_price"])),
        ("Price / Unit", _cur(derived["price_per_unit"])),
        ("Current NOI", _cur(deal["current_noi"])),
        ("Occupancy", f"{deal['occupancy']*100:.0f}%"),
    ])

    _subsection(pdf, "Return Metrics")
    _kv_table(pdf, [
        ("Levered IRR", _pct(m.get("levered_irr"))),
        ("After-Tax IRR", _pct(m.get("after_tax_irr"))),
        ("Equity Multiple", f"{m['equity_multiple']:.2f}x" if m.get("equity_multiple") else "N/A"),
        ("Cash-on-Cash (Yr 1)", _pct(m.get("cash_on_cash_yr1"))),
        ("DSCR (Yr 1)", f"{m['dscr_yr1']:.2f}x" if m.get("dscr_yr1") else "N/A"),
        ("Going-In Cap", _pct(m.get("going_in_cap_rate"))),
        ("Exit Cap", _pct(m.get("exit_cap_rate"))),
    ])

    pdf.set_font("Helvetica", "B", 14)
    if rec == "BUY":
        pdf.set_text_color(0, 128, 0)
    elif "HOLD" in rec:
        pdf.set_text_color(255, 153, 0)
    else:
        pdf.set_text_color(204, 0, 0)
    pdf.cell(0, 10, f"Recommendation: {rec}", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(*BLACK)

    # ===== SOURCES & USES =====
    pdf.add_page()
    _section(pdf, "2. Capital Structure")

    total_s = su["sources"]["total"]
    _subsection(pdf, "Sources")
    _data_table(pdf, ["Source", "Amount", "% of Total"], [
        ["Senior Debt", _cur(su["sources"]["senior_debt"]), _pct(su["sources"]["senior_debt"]/total_s) if total_s else "0%"],
        ["Sponsor Equity", _cur(su["sources"]["sponsor_equity"]), _pct(su["sources"]["sponsor_equity"]/total_s) if total_s else "0%"],
        ["Total", _cur(total_s), "100.00%"],
    ])

    _subsection(pdf, "Uses")
    total_u = su["uses"]["total"]
    uses_rows = [
        ["Purchase Price", _cur(su["uses"]["purchase_price"]), _pct(su["uses"]["purchase_price"]/total_u) if total_u else "0%"],
        ["Closing Costs", _cur(su["uses"]["closing_costs"]), _pct(su["uses"]["closing_costs"]/total_u) if total_u else "0%"],
    ]
    if su["uses"]["deferred_maintenance"] > 0:
        uses_rows.append(["Deferred Maint.", _cur(su["uses"]["deferred_maintenance"]), _pct(su["uses"]["deferred_maintenance"]/total_u)])
    if su["uses"]["planned_capex"] > 0:
        uses_rows.append(["Planned CapEx", _cur(su["uses"]["planned_capex"]), _pct(su["uses"]["planned_capex"]/total_u)])
    uses_rows.append(["Total", _cur(total_u), "100.00%"])
    _data_table(pdf, ["Use", "Amount", "% of Total"], uses_rows)

    # ===== PRO FORMA =====
    _section(pdf, "3. Pro Forma Summary")
    hold = deal["hold_period_years"]
    pf_headers = ["Year", "Rent/Unit", "Occ", "NOI", "BTCF"]
    pf_rows = []
    for yr in pf[:hold]:
        pf_rows.append([
            f"Yr {yr['year']}", f"${yr['rent_per_unit']:,.0f}",
            f"{yr['occupancy']*100:.0f}%", _cur(yr["noi"]), _cur(yr["btcf"]),
        ])
    _data_table(pdf, pf_headers, pf_rows)

    # ===== EXIT SUMMARY =====
    _subsection(pdf, "Exit Summary")
    _kv_table(pdf, [
        ("Exit Year", str(rev["exit_year"])),
        ("Forward NOI", _cur(rev["forward_noi"])),
        ("Exit Cap Rate", _pct(rev["exit_cap_rate"])),
        ("Gross Sale Price", _cur(rev["sale_price"])),
        ("Net Sale Proceeds", _cur(rev["net_sale_proceeds"])),
    ])

    # ===== SENSITIVITY =====
    if sensitivity and sensitivity.get("exit_cap"):
        pdf.add_page()
        _section(pdf, "4. Sensitivity Analysis")
        _subsection(pdf, "Exit Cap Rate")
        sens_rows = []
        for row in sensitivity["exit_cap"]:
            irr_str = f"{row['irr']:.1f}%" if row["irr"] is not None else "N/A"
            sens_rows.append([
                f"{row['exit_cap_rate']:.2f}%", _cur(row["sale_price"]),
                irr_str, f"{row['equity_multiple']:.2f}x"
            ])
        _data_table(pdf, ["Exit Cap", "Sale Price", "IRR", "EM"], sens_rows)

    # ===== MONTE CARLO =====
    if has_mc:
        pdf.add_page()
        _section(pdf, "5. Monte Carlo Simulation")
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(*BLACK)
        pdf.multi_cell(0, 6, monte_carlo.get("summary", ""))
        pdf.ln(5)

        _subsection(pdf, "IRR Percentiles")
        pctile_rows = [[k, f"{v:.1f}%"] for k, v in monte_carlo.get("percentiles", {}).items()]
        _data_table(pdf, ["Percentile", "IRR"], pctile_rows, col_widths=[90, 90])

        _subsection(pdf, "Probabilities")
        prob_rows = [[k, f"{v:.1f}%"] for k, v in monte_carlo.get("probabilities", {}).items()]
        _data_table(pdf, ["Threshold", "Probability"], prob_rows, col_widths=[90, 90])

    # ===== ML VALUATION =====
    if has_ml:
        pdf.add_page()
        _section(pdf, "ML-Based Property Valuation")
        _kv_table(pdf, [
            ("Assessment", ml_valuation["assessment"]),
            ("Predicted Value/Unit", _cur(ml_valuation["predicted_value_per_unit"])),
            ("Predicted Total", _cur(ml_valuation["predicted_total_value"])),
            ("Actual Price/Unit", _cur(ml_valuation["actual_price_per_unit"])),
            ("Premium/Discount", f"{ml_valuation['premium_discount_pct']:.1f}%"),
            ("Test R²", f"{ml_valuation.get('test_r2', 'N/A')}"),
            ("Test MAPE", f"{ml_valuation.get('test_mape', 'N/A')}%"),
            ("Features Used", str(ml_valuation["features_used"])),
        ])

    # ===== RENT FORECAST =====
    if has_rent:
        pdf.add_page()
        _section(pdf, "Predictive Rent Growth Model")
        _kv_table(pdf, [
            ("Method", rent_prediction["method"]),
            ("Data Source", rent_prediction["data_source"]),
            ("Avg Predicted Growth", f"{rent_prediction['avg_predicted_growth']}%"),
            ("Historical Avg Growth", f"{rent_prediction['historical_avg']}%"),
        ])

        _subsection(pdf, "Forecast")
        forecast_rows = []
        for i in range(len(rent_prediction["predicted_rates"])):
            forecast_rows.append([
                f"Year {i+1}",
                f"{rent_prediction['predicted_rates'][i]:.2f}%",
                f"${rent_prediction['predicted_rents_per_unit'][i]:,.0f}",
                _cur(rent_prediction["predicted_annual_revenue"][i]),
            ])
        _data_table(pdf, ["Year", "Growth", "Rent/Unit", "Revenue"], forecast_rows)

    # ===== BACKTEST =====
    if has_bt:
        _section(pdf, "Model Backtest Validation")
        _kv_table(pdf, [
            ("MAE", f"{backtest['mae']:.3f}%"),
            ("RMSE", f"{backtest['rmse']:.3f}%"),
            ("Direction Accuracy", f"{backtest['direction_accuracy']}%" if backtest.get("direction_accuracy") else "N/A"),
            ("Quality", backtest.get("quality", "N/A")),
        ])

    # ===== LEASE ANALYSIS =====
    if has_lease:
        pdf.add_page()
        _section(pdf, "Lease Document Analysis")
        if lease_analysis.get("summary"):
            pdf.set_font("Helvetica", "I", 10)
            pdf.set_text_color(*BLACK)
            pdf.multi_cell(0, 6, lease_analysis["summary"])
            pdf.ln(3)

        lease_items = []
        for label, key in [("Tenant", "tenant_name"), ("Lease Type", "lease_type"),
                           ("Monthly Rent", "monthly_rent"), ("Term (mo)", "lease_term_months"),
                           ("Escalation", "escalation_clause")]:
            val = lease_analysis.get(key)
            if val is not None:
                if isinstance(val, (int, float)) and val > 100:
                    lease_items.append((label, _cur(val)))
                else:
                    lease_items.append((label, str(val)))
        if lease_items:
            _kv_table(pdf, lease_items)

    # ===== DISCLAIMER =====
    pdf.add_page()
    _section(pdf, "Disclaimers")
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(100, 100, 100)
    pdf.multi_cell(0, 5, (
        "This investment memorandum is provided for informational purposes only and does not "
        "constitute an offer to sell or a solicitation of an offer to buy any security. Financial "
        "projections are based on assumptions that may not be realized. Actual results may differ "
        "materially. Past performance is not indicative of future results. Investors should conduct "
        "their own due diligence and consult with their advisors before making any investment decision."
    ))
    pdf.ln(5)
    if has_ml or has_lease or has_rent:
        pdf.multi_cell(0, 5, (
            "AI/ML Disclaimer: Machine learning valuations, NLP lease analysis, and predictive rent "
            "models use synthetic training data and statistical methods that carry inherent limitations. "
            "These outputs should supplement, not replace, professional judgment and independent appraisals."
        ))

    # Save
    filename = f"investment_report_{job_id}.pdf"
    filepath = os.path.join(OUTPUT_DIR, filename)
    pdf.output(filepath)
    logger.info(f"PDF report generated: {filepath}")
    return filepath
