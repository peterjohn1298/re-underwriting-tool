"""Generate institutional-grade Excel workbook with 8 tabs and charts."""

import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.utils import get_column_letter

from config import OUTPUT_DIR

# --- Style Constants ---
NAVY = "1B2A4A"
DARK_BLUE = "2C3E6B"
GOLD = "C8A951"
WHITE = "FFFFFF"
LIGHT_GRAY = "F2F2F2"
MED_GRAY = "D9D9D9"

HEADER_FONT = Font(name="Calibri", bold=True, size=11, color=WHITE)
HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type="solid")
SUBHEADER_FONT = Font(name="Calibri", bold=True, size=11, color=NAVY)
SUBHEADER_FILL = PatternFill(start_color=MED_GRAY, end_color=MED_GRAY, fill_type="solid")
NORMAL_FONT = Font(name="Calibri", size=11)
TITLE_FONT = Font(name="Calibri", bold=True, size=14, color=NAVY)
ALT_FILL = PatternFill(start_color=LIGHT_GRAY, end_color=LIGHT_GRAY, fill_type="solid")
CURRENCY_FMT = '$#,##0'
PCT_FMT = '0.00%'
THIN_BORDER = Border(
    left=Side(style="thin", color=MED_GRAY),
    right=Side(style="thin", color=MED_GRAY),
    top=Side(style="thin", color=MED_GRAY),
    bottom=Side(style="thin", color=MED_GRAY),
)
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")


def _hdr_row(ws, row, max_col):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def _data_row(ws, row, max_col, alt=False):
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = NORMAL_FONT
        cell.border = THIN_BORDER
        if alt:
            cell.fill = ALT_FILL


def _title(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.font = TITLE_FONT


def _widths(ws, max_col, w=15):
    for col in range(1, max_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = w


def generate_excel(pro_forma: dict, market_data: dict, job_id: str) -> str:
    wb = Workbook()
    inp = pro_forma["inputs"]
    deal = inp["deal"]
    derived = inp["derived"]
    metrics = pro_forma["metrics"]
    pf = pro_forma["pro_forma"]
    su = pro_forma["sources_uses"]
    amort = pro_forma["amortization_annual"]
    rev = pro_forma["reversion"]

    # ========== TAB 1: Summary ==========
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_properties.tabColor = NAVY

    _title(ws, 1, 1, "Investment Summary")
    r = 3
    snap = [
        ("Property", derived["property_name"]),
        ("Type", deal["property_type"]),
        ("Address", deal["address"]),
        ("Year Built", deal["year_built"]),
        ("Units / SF", f"{deal['total_units']} units / {deal['total_sf']:,.0f} SF"),
        ("Purchase Price", deal["purchase_price"]),
        ("Price / Unit", derived["price_per_unit"]),
        ("Current NOI", deal["current_noi"]),
        ("In-Place Rent", f"${deal['in_place_rent']:,.0f}/mo"),
        ("Market Rent", f"${deal['market_rent']:,.0f}/mo"),
        ("Occupancy", deal["occupancy"]),
        ("Total CapEx", derived["total_capex"]),
    ]
    for label, val in snap:
        ws.cell(row=r, column=1, value=label).font = SUBHEADER_FONT
        c = ws.cell(row=r, column=2, value=val)
        c.font = NORMAL_FONT
        if isinstance(val, float) and val < 1 and val > 0:
            c.number_format = PCT_FMT
        elif isinstance(val, (int, float)) and val > 1000:
            c.number_format = CURRENCY_FMT
        r += 1

    r += 1
    _title(ws, r, 1, "Return Metrics")
    r += 1
    for c_idx, h in enumerate(["Metric", "Value"], 1):
        ws.cell(row=r, column=c_idx, value=h)
    _hdr_row(ws, r, 2)
    r += 1

    ret_items = [
        ("Levered IRR", metrics["levered_irr"], True),
        ("Unlevered IRR", metrics["unlevered_irr"], True),
        ("Equity Multiple", metrics["equity_multiple"], False),
        ("Cash-on-Cash (Yr 1)", metrics["cash_on_cash_yr1"], True),
        ("DSCR (Yr 1)", metrics["dscr_yr1"], False),
        ("Stabilized DSCR", metrics["stabilized_dscr"], False),
        ("Yield on Cost", metrics["yield_on_cost"], True),
        ("Stabilized YOC", metrics["stabilized_yoc"], True),
        ("Going-In Cap Rate", metrics["going_in_cap_rate"], True),
        ("Exit Cap Rate", metrics["exit_cap_rate"], True),
    ]
    for label, val, is_pct in ret_items:
        ws.cell(row=r, column=1, value=label).font = NORMAL_FONT
        c = ws.cell(row=r, column=2)
        if val is not None and is_pct:
            c.value = val
            c.number_format = PCT_FMT
        elif val is not None:
            c.value = round(val, 2)
            c.number_format = '0.00"x"'
        else:
            c.value = "N/A"
        c.font = NORMAL_FONT
        r += 1

    # NOI chart
    if pf:
        r += 2
        ws.cell(row=r, column=1, value="Year")
        ws.cell(row=r, column=2, value="NOI")
        for i, yr in enumerate(pf):
            ws.cell(row=r+1+i, column=1, value=f"Year {yr['year']}")
            ws.cell(row=r+1+i, column=2, value=yr["noi"])
        chart = BarChart()
        chart.title = "Net Operating Income"
        chart.style = 10
        chart.width = 20
        chart.height = 12
        chart.add_data(Reference(ws, min_col=2, min_row=r, max_row=r+len(pf)), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=r+1, max_row=r+len(pf)))
        chart.series[0].graphicalProperties.solidFill = NAVY
        ws.add_chart(chart, f"D3")
    _widths(ws, 5)

    # ========== TAB 2: Sources & Uses ==========
    ws2 = wb.create_sheet("Sources & Uses")
    ws2.sheet_properties.tabColor = DARK_BLUE
    _title(ws2, 1, 1, "Sources & Uses of Funds")

    r = 3
    for h_idx, h in enumerate(["SOURCES", "Amount", "% of Total"], 1):
        ws2.cell(row=r, column=h_idx, value=h)
    _hdr_row(ws2, r, 3)
    r += 1
    total_s = su["sources"]["total"]
    for label, key in [("Senior Debt", "senior_debt"), ("Sponsor Equity", "sponsor_equity")]:
        ws2.cell(row=r, column=1, value=label).font = NORMAL_FONT
        ws2.cell(row=r, column=2, value=su["sources"][key]).number_format = CURRENCY_FMT
        ws2.cell(row=r, column=3, value=su["sources"][key]/total_s if total_s else 0).number_format = PCT_FMT
        r += 1
    ws2.cell(row=r, column=1, value="Total Sources").font = SUBHEADER_FONT
    ws2.cell(row=r, column=2, value=total_s).number_format = CURRENCY_FMT
    ws2.cell(row=r, column=2).font = SUBHEADER_FONT

    r += 2
    for h_idx, h in enumerate(["USES", "Amount", "% of Total"], 1):
        ws2.cell(row=r, column=h_idx, value=h)
    _hdr_row(ws2, r, 3)
    r += 1
    total_u = su["uses"]["total"]
    uses_items = [
        ("Purchase Price", su["uses"]["purchase_price"]),
        ("Closing Costs", su["uses"]["closing_costs"]),
    ]
    if su["uses"]["deferred_maintenance"] > 0:
        uses_items.append(("Deferred Maintenance", su["uses"]["deferred_maintenance"]))
    if su["uses"]["planned_capex"] > 0:
        desc = deal.get("capex_description", "")
        label = f"Planned CapEx ({desc})" if desc else "Planned CapEx"
        uses_items.append((label, su["uses"]["planned_capex"]))

    for label, val in uses_items:
        ws2.cell(row=r, column=1, value=label).font = NORMAL_FONT
        ws2.cell(row=r, column=2, value=val).number_format = CURRENCY_FMT
        ws2.cell(row=r, column=3, value=val/total_u if total_u else 0).number_format = PCT_FMT
        r += 1
    ws2.cell(row=r, column=1, value="Total Uses").font = SUBHEADER_FONT
    ws2.cell(row=r, column=2, value=total_u).number_format = CURRENCY_FMT
    ws2.cell(row=r, column=2).font = SUBHEADER_FONT
    _widths(ws2, 3)

    # ========== TAB 3: Pro Forma ==========
    ws3 = wb.create_sheet("Pro Forma")
    ws3.sheet_properties.tabColor = DARK_BLUE
    _title(ws3, 1, 1, "10-Year Pro Forma")

    headers = [""] + [f"Year {i}" for i in range(1, 11)]
    r = 3
    for c, h in enumerate(headers, 1):
        ws3.cell(row=r, column=c, value=h)
    _hdr_row(ws3, r, len(headers))

    rows = [
        ("Rent / Unit ($/mo)", "rent_per_unit", CURRENCY_FMT),
        ("Occupancy", "occupancy", PCT_FMT),
        ("Gross Potential Rent", "gross_potential_rent", CURRENCY_FMT),
        ("Less: Vacancy", "vacancy_loss", CURRENCY_FMT),
        ("Other Income", "other_income", CURRENCY_FMT),
        ("Effective Gross Income", "effective_gross_income", CURRENCY_FMT),
        ("", None, None),
        ("Management Fee", "management_fee", CURRENCY_FMT),
        ("Property Tax", "property_tax", CURRENCY_FMT),
        ("Insurance", "insurance", CURRENCY_FMT),
        ("Utilities", "utilities", CURRENCY_FMT),
        ("Repairs & Maintenance", "repairs_maintenance", CURRENCY_FMT),
        ("General & Admin", "general_admin", CURRENCY_FMT),
        ("Other Expenses", "other_expenses", CURRENCY_FMT),
        ("Replacement Reserves", "replacement_reserves", CURRENCY_FMT),
        ("Total Expenses", "total_expenses", CURRENCY_FMT),
        ("", None, None),
        ("Net Operating Income", "noi", CURRENCY_FMT),
        ("Debt Service", "debt_service", CURRENCY_FMT),
        ("Before-Tax Cash Flow", "btcf", CURRENCY_FMT),
    ]
    bold_rows = {"Effective Gross Income", "Total Expenses", "Net Operating Income", "Before-Tax Cash Flow"}

    for i, (label, key, fmt) in enumerate(rows):
        r += 1
        ws3.cell(row=r, column=1, value=label)
        if label in bold_rows:
            ws3.cell(row=r, column=1).font = SUBHEADER_FONT
        else:
            ws3.cell(row=r, column=1).font = NORMAL_FONT
        if key:
            for yr_idx, yr_data in enumerate(pf):
                c = ws3.cell(row=r, column=yr_idx + 2, value=yr_data[key])
                if fmt:
                    c.number_format = fmt
                c.font = NORMAL_FONT
        _data_row(ws3, r, len(headers), alt=(i % 2 == 0))

    _widths(ws3, len(headers), 14)
    ws3.column_dimensions["A"].width = 24

    # ========== TAB 4: Loan ==========
    ws4 = wb.create_sheet("Loan")
    ws4.sheet_properties.tabColor = DARK_BLUE
    _title(ws4, 1, 1, "Loan Amortization")

    r = 3
    for label, val, fmt in [
        ("Loan Amount", derived["loan_amount"], CURRENCY_FMT),
        ("LTV", deal["ltv"], PCT_FMT),
        ("Interest Rate", deal["interest_rate"], PCT_FMT),
        ("Amortization", f"{deal['amortization_years']} years", None),
        ("Loan Term", f"{deal['loan_term_years']} years", None),
        ("IO Period", f"{deal['io_period_years']} years", None),
    ]:
        ws4.cell(row=r, column=1, value=label).font = SUBHEADER_FONT
        c = ws4.cell(row=r, column=2, value=val)
        if fmt:
            c.number_format = fmt
        c.font = NORMAL_FONT
        r += 1

    if amort:
        r += 1
        hdrs = ["Year", "Total Payment", "Principal", "Interest", "Ending Balance"]
        for c_idx, h in enumerate(hdrs, 1):
            ws4.cell(row=r, column=c_idx, value=h)
        _hdr_row(ws4, r, len(hdrs))
        r += 1
        for i, yr in enumerate(amort):
            ws4.cell(row=r, column=1, value=yr["year"]).font = NORMAL_FONT
            for c_idx, key in enumerate(["total_payment", "total_principal", "total_interest", "ending_balance"], 2):
                ws4.cell(row=r, column=c_idx, value=yr[key]).number_format = CURRENCY_FMT
            _data_row(ws4, r, len(hdrs), alt=(i % 2 == 0))
            r += 1

        chart = LineChart()
        chart.title = "Loan Balance"
        chart.style = 10
        chart.width = 18
        chart.height = 10
        ds = r - len(amort)
        chart.add_data(Reference(ws4, min_col=5, min_row=ds-1, max_row=r-1), titles_from_data=True)
        chart.set_categories(Reference(ws4, min_col=1, min_row=ds, max_row=r-1))
        ws4.add_chart(chart, f"A{r+1}")
    _widths(ws4, 5)

    # ========== TAB 5: Comps ==========
    ws5 = wb.create_sheet("Comps")
    ws5.sheet_properties.tabColor = DARK_BLUE
    _title(ws5, 1, 1, "Comparable Sales")

    comps = market_data.get("comps", {}).get("comps", [])
    r = 3
    comp_h = ["#", "Property", "Price/Unit", "Cap Rate", "Year", "Units"]
    for c_idx, h in enumerate(comp_h, 1):
        ws5.cell(row=r, column=c_idx, value=h)
    _hdr_row(ws5, r, len(comp_h))
    r += 1

    for i, comp in enumerate(comps[:7], 1):
        if not isinstance(comp, dict):
            continue
        ws5.cell(row=r, column=1, value=i).font = NORMAL_FONT
        ws5.cell(row=r, column=2, value=comp.get("name", comp.get("source", f"Comp {i}"))).font = NORMAL_FONT
        ppu = comp.get("price_per_unit", "")
        if isinstance(ppu, (int, float)):
            ws5.cell(row=r, column=3, value=ppu).number_format = CURRENCY_FMT
        else:
            ws5.cell(row=r, column=3, value=ppu)
        cap = comp.get("cap_rate", "")
        if isinstance(cap, (int, float)):
            ws5.cell(row=r, column=4, value=cap/100 if cap > 1 else cap).number_format = PCT_FMT
        else:
            ws5.cell(row=r, column=4, value=cap)
        ws5.cell(row=r, column=5, value=comp.get("year", "")).font = NORMAL_FONT
        ws5.cell(row=r, column=6, value=comp.get("units", "")).font = NORMAL_FONT
        _data_row(ws5, r, len(comp_h), alt=(i % 2 == 0))
        r += 1

    # Stats
    ppus = [c["price_per_unit"] for c in comps if isinstance(c, dict) and isinstance(c.get("price_per_unit"), (int, float))]
    if ppus:
        r += 1
        ws5.cell(row=r, column=2, value="Average").font = SUBHEADER_FONT
        ws5.cell(row=r, column=3, value=round(sum(ppus)/len(ppus))).number_format = CURRENCY_FMT
    _widths(ws5, len(comp_h))

    # ========== TAB 6: Sensitivity ==========
    ws6 = wb.create_sheet("Sensitivity")
    ws6.sheet_properties.tabColor = DARK_BLUE
    _title(ws6, 1, 1, "Sensitivity Analysis")

    from models.metrics import sensitivity_table_exit_cap
    exit_sens = sensitivity_table_exit_cap(
        base_noi_at_exit=rev["forward_noi"],
        base_exit_cap=rev["exit_cap_rate"],
        equity_invested=derived["equity_required"],
        annual_cfs=pro_forma["annual_btcfs"][:rev["exit_year"]],
        sale_costs_pct=deal["sale_costs_pct"],
        loan_balance_at_exit=rev["loan_balance"],
    )

    r = 3
    ws6.cell(row=r, column=1, value="Exit Cap Rate Sensitivity").font = SUBHEADER_FONT
    r += 1
    sh = ["Exit Cap", "Sale Price", "IRR", "Equity Multiple"]
    for c_idx, h in enumerate(sh, 1):
        ws6.cell(row=r, column=c_idx, value=h)
    _hdr_row(ws6, r, len(sh))
    r += 1

    for i, row_data in enumerate(exit_sens):
        ws6.cell(row=r, column=1, value=row_data["exit_cap_rate"]/100).number_format = PCT_FMT
        ws6.cell(row=r, column=2, value=row_data["sale_price"]).number_format = CURRENCY_FMT
        irr_c = ws6.cell(row=r, column=3)
        if row_data["irr"] is not None:
            irr_c.value = row_data["irr"] / 100
            irr_c.number_format = PCT_FMT
            if row_data["irr"] >= 15:
                irr_c.fill = GREEN_FILL
            elif row_data["irr"] >= 10:
                irr_c.fill = YELLOW_FILL
            else:
                irr_c.fill = RED_FILL
        else:
            irr_c.value = "N/A"
        ws6.cell(row=r, column=4, value=row_data["equity_multiple"])
        _data_row(ws6, r, len(sh), alt=(i % 2 == 0))
        r += 1
    _widths(ws6, 4)

    # ========== TAB 7: Market ==========
    ws7 = wb.create_sheet("Market")
    ws7.sheet_properties.tabColor = DARK_BLUE
    _title(ws7, 1, 1, "Market Overview")

    r = 3
    demo = market_data.get("demographics", {})
    ws7.cell(row=r, column=1, value="Location").font = SUBHEADER_FONT
    ws7.cell(row=r, column=2, value=f"{derived['city']}, {derived['state']}").font = NORMAL_FONT
    r += 2

    summary = demo.get("summary", "Market data not available.")
    c = ws7.cell(row=r, column=1, value=summary)
    c.font = NORMAL_FONT
    c.alignment = Alignment(wrap_text=True)
    ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    r += 2

    cap_data = market_data.get("cap_rates", {})
    ws7.cell(row=r, column=1, value="Market Cap Rate").font = NORMAL_FONT
    avg_cap = cap_data.get("average_cap_rate", "N/A")
    if isinstance(avg_cap, (int, float)):
        ws7.cell(row=r, column=2, value=avg_cap/100).number_format = PCT_FMT
    else:
        ws7.cell(row=r, column=2, value=avg_cap)
    r += 1

    rent_data = market_data.get("rent_trends", {})
    ws7.cell(row=r, column=1, value="Avg Rent Growth").font = NORMAL_FONT
    avg_rent = rent_data.get("average_growth", "N/A")
    if isinstance(avg_rent, (int, float)):
        ws7.cell(row=r, column=2, value=avg_rent/100).number_format = PCT_FMT
    else:
        ws7.cell(row=r, column=2, value=avg_rent)

    # NOI line chart
    if pf:
        r += 2
        ws7.cell(row=r, column=1, value="Year")
        ws7.cell(row=r, column=2, value="NOI")
        for i, yr in enumerate(pf):
            ws7.cell(row=r+1+i, column=1, value=f"Yr {yr['year']}")
            ws7.cell(row=r+1+i, column=2, value=yr["noi"])
        chart = LineChart()
        chart.title = "NOI Growth Trajectory"
        chart.style = 10
        chart.width = 18
        chart.height = 10
        chart.add_data(Reference(ws7, min_col=2, min_row=r, max_row=r+len(pf)), titles_from_data=True)
        chart.set_categories(Reference(ws7, min_col=1, min_row=r+1, max_row=r+len(pf)))
        ws7.add_chart(chart, f"D{r}")
    _widths(ws7, 4)

    # ========== TAB 8: Assumptions ==========
    ws8 = wb.create_sheet("Assumptions")
    ws8.sheet_properties.tabColor = DARK_BLUE
    _title(ws8, 1, 1, "Underwriting Assumptions")

    r = 3
    sections = [
        ("Deal Inputs", [
            ("Property Type", deal["property_type"], None),
            ("Address", deal["address"], None),
            ("Year Built", deal["year_built"], None),
            ("Units", deal["total_units"], None),
            ("Total SF", deal["total_sf"], '#,##0'),
            ("In-Place Rent", deal["in_place_rent"], CURRENCY_FMT),
            ("Market Rent", deal["market_rent"], CURRENCY_FMT),
            ("Occupancy", deal["occupancy"], PCT_FMT),
            ("Current NOI", deal["current_noi"], CURRENCY_FMT),
        ]),
        ("Acquisition", [
            ("Purchase Price", deal["purchase_price"], CURRENCY_FMT),
            ("Closing Costs", deal["closing_costs_pct"], PCT_FMT),
            ("Deferred Maintenance", deal["deferred_maintenance"], CURRENCY_FMT),
            ("Planned CapEx", deal["planned_capex"], CURRENCY_FMT),
            ("Total Project Cost", derived["total_project_cost"], CURRENCY_FMT),
        ]),
        ("Financing", [
            ("LTV", deal["ltv"], PCT_FMT),
            ("Interest Rate", deal["interest_rate"], PCT_FMT),
            ("Amortization", f"{deal['amortization_years']} years", None),
            ("Loan Term", f"{deal['loan_term_years']} years", None),
            ("IO Period", f"{deal['io_period_years']} years", None),
            ("Loan Amount", derived["loan_amount"], CURRENCY_FMT),
            ("Equity Required", derived["equity_required"], CURRENCY_FMT),
        ]),
        ("Operating", [
            ("Revenue Growth", deal["revenue_growth_rate"], PCT_FMT),
            ("Expense Growth", deal["expense_growth_rate"], PCT_FMT),
            ("Management Fee", deal["management_fee_pct"], PCT_FMT),
            ("Reserves/Unit", deal["replacement_reserves_per_unit"], CURRENCY_FMT),
            ("Expense Ratio", derived["expense_ratio"], PCT_FMT),
        ]),
        ("Exit", [
            ("Hold Period", f"{deal['hold_period_years']} years", None),
            ("Exit Cap Spread", deal["exit_cap_rate_spread"], PCT_FMT),
            ("Sale Costs", deal["sale_costs_pct"], PCT_FMT),
            ("Going-In Cap Rate", derived["going_in_cap_rate"], PCT_FMT),
            ("Exit Cap Rate", derived["exit_cap_rate"], PCT_FMT),
        ]),
    ]

    for section_name, items in sections:
        ws8.cell(row=r, column=1, value=section_name).font = HEADER_FONT
        ws8.cell(row=r, column=1).fill = HEADER_FILL
        ws8.cell(row=r, column=2).fill = HEADER_FILL
        r += 1
        for label, val, fmt in items:
            ws8.cell(row=r, column=1, value=label).font = NORMAL_FONT
            c = ws8.cell(row=r, column=2, value=val)
            c.font = NORMAL_FONT
            if fmt:
                c.number_format = fmt
            r += 1
        r += 1
    _widths(ws8, 2, 20)

    # Save
    filename = f"underwriting_{job_id}.xlsx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    wb.save(filepath)
    return filepath
