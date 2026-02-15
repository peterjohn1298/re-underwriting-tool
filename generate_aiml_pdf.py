"""Generate a detailed PDF explaining the AI/ML integration in the RE Underwriting Tool."""

from fpdf import FPDF
from datetime import datetime


class PDF(FPDF):
    NAVY = (27, 42, 74)
    GOLD = (200, 169, 81)
    DARK = (40, 40, 40)
    GRAY = (100, 100, 100)
    WHITE = (255, 255, 255)
    LIGHT_BG = (245, 245, 250)

    def header(self):
        if self.page_no() > 1:
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(*self.GRAY)
            self.cell(0, 8, "RE Underwriting Tool - AI/ML Integration Technical Document", align="R")
            self.ln(4)
            self.set_draw_color(*self.NAVY)
            self.set_line_width(0.3)
            self.line(10, self.get_y(), 200, self.get_y())
            self.ln(6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(*self.GRAY)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")

    def section_title(self, title):
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(*self.NAVY)
        self.cell(0, 12, title, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(*self.GOLD)
        self.set_line_width(0.8)
        self.line(10, self.get_y(), 80, self.get_y())
        self.ln(6)

    def sub_title(self, title):
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(*self.NAVY)
        self.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")
        self.ln(2)

    def sub_sub_title(self, title):
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(80, 80, 80)
        self.cell(0, 8, title, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def body_text(self, text):
        self.set_font("Helvetica", "", 10)
        self.set_text_color(*self.DARK)
        self.multi_cell(0, 5.5, text)
        self.ln(3)

    def bullet(self, text):
        self.set_font("Helvetica", "", 10)
        self.set_text_color(*self.DARK)
        x = self.get_x()
        self.cell(6, 5.5, "-")
        self.multi_cell(0, 5.5, text)
        self.ln(1)

    def code_block(self, text):
        self.set_fill_color(*self.LIGHT_BG)
        self.set_font("Courier", "", 9)
        self.set_text_color(60, 60, 60)
        y = self.get_y()
        self.multi_cell(0, 5, text, fill=True)
        self.ln(4)

    def callout_box(self, text, color=None):
        if color is None:
            color = self.NAVY
        self.set_fill_color(color[0], color[1], color[2])
        self.set_text_color(*self.WHITE)
        self.set_font("Helvetica", "B", 10)
        x, y = self.get_x(), self.get_y()
        self.rect(10, y, 190, 10, style="F")
        self.set_xy(14, y + 2)
        self.multi_cell(182, 5.5, text)
        self.set_y(y + 14)
        self.ln(2)

    def table(self, headers, rows, col_widths=None):
        if col_widths is None:
            col_widths = [190 / len(headers)] * len(headers)

        # Header
        self.set_fill_color(*self.NAVY)
        self.set_text_color(*self.WHITE)
        self.set_font("Helvetica", "B", 9)
        for i, h in enumerate(headers):
            self.cell(col_widths[i], 7, h, border=1, fill=True, align="C")
        self.ln()

        # Rows
        self.set_font("Helvetica", "", 9)
        self.set_text_color(*self.DARK)
        for ri, row in enumerate(rows):
            if ri % 2 == 0:
                self.set_fill_color(245, 245, 245)
            else:
                self.set_fill_color(255, 255, 255)
            for i, val in enumerate(row):
                align = "L" if i == 0 else "C"
                self.cell(col_widths[i], 6.5, str(val), border=1, fill=True, align=align)
            self.ln()
        self.ln(4)


def build_pdf():
    pdf = PDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)

    # ===== COVER PAGE =====
    pdf.add_page()
    pdf.ln(50)
    pdf.set_font("Helvetica", "B", 32)
    pdf.set_text_color(*PDF.NAVY)
    pdf.cell(0, 15, "AI/ML Integration", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 15, "Technical Document", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(8)
    pdf.set_draw_color(*PDF.GOLD)
    pdf.set_line_width(1.5)
    pdf.line(60, pdf.get_y(), 150, pdf.get_y())
    pdf.ln(10)
    pdf.set_font("Helvetica", "", 16)
    pdf.set_text_color(*PDF.GOLD)
    pdf.cell(0, 10, "Real Estate Investment Underwriting Tool", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)
    pdf.set_font("Helvetica", "", 12)
    pdf.set_text_color(*PDF.GRAY)
    pdf.cell(0, 8, "How Machine Learning, Natural Language Processing,", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, "and Predictive Analytics Integrate Into", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, "Commercial Real Estate Financial Modeling", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(30)
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(*PDF.NAVY)
    pdf.cell(0, 8, datetime.now().strftime("%B %d, %Y"), align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(200, 0, 0)
    pdf.cell(0, 8, "CONFIDENTIAL", align="C", new_x="LMARGIN", new_y="NEXT")

    # ===== TABLE OF CONTENTS =====
    pdf.add_page()
    pdf.section_title("Table of Contents")
    pdf.ln(4)
    toc = [
        ("1.", "What This Tool Does", "3"),
        ("2.", "The Three AI/ML Systems  -- Overview", "4"),
        ("3.", "System 1: Predictive Rent Growth Model", "5"),
        ("4.", "System 2: ML Property Valuation", "8"),
        ("5.", "System 3: Lease NLP Analysis", "11"),
        ("6.", "The Integrated Recommendation Engine", "14"),
        ("7.", "Complete Pipeline  -- How It All Flows Together", "16"),
        ("8.", "Technology Stack", "17"),
        ("9.", "Data Sources and API Integrations", "18"),
        ("10.", "Academic Honesty and Limitations", "19"),
    ]
    for num, title, page in toc:
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(*PDF.NAVY)
        pdf.cell(10, 8, num)
        pdf.set_font("Helvetica", "", 11)
        pdf.set_text_color(*PDF.DARK)
        pdf.cell(150, 8, title)
        pdf.set_font("Helvetica", "", 11)
        pdf.set_text_color(*PDF.GRAY)
        pdf.cell(0, 8, page, align="R", new_x="LMARGIN", new_y="NEXT")

    # ===== SECTION 1: WHAT THIS TOOL DOES =====
    pdf.add_page()
    pdf.section_title("1. What This Tool Does")

    pdf.body_text(
        "Imagine you want to buy an apartment building for $12.5 million. Before you spend that money, you need to "
        "answer one question: \"Is this a good deal?\""
    )
    pdf.body_text(
        "That is what underwriting does. It is a financial analysis that projects what the property will earn over "
        "7-10 years, what it will be worth when you sell it, and whether the returns justify the risk."
    )
    pdf.body_text(
        "Traditionally, an analyst opens Excel, manually types in assumptions (rent, expenses, growth rates), and "
        "builds a spreadsheet. This takes hours and is only as good as the assumptions the analyst guesses."
    )
    pdf.body_text(
        "Our tool automates the entire process. But it goes further by adding three AI/ML systems that make the "
        "analysis smarter, more data-driven, and more transparent about its limitations."
    )

    pdf.sub_title("Key Terms for Non-Finance Readers")
    terms = [
        ("NOI (Net Operating Income)", "Revenue minus operating expenses. How much the building earns before paying the mortgage."),
        ("Cap Rate", "NOI divided by property price. Think of it as the 'yield' on a real estate investment, like an interest rate."),
        ("IRR (Internal Rate of Return)", "The annualized return on your investment, accounting for the timing of all cash flows. The single most important metric."),
        ("Pro Forma", "A 10-year financial projection. Year-by-year forecast of income, expenses, and cash flows."),
        ("DSCR (Debt Service Coverage Ratio)", "NOI divided by mortgage payments. Must be above 1.0x or you cannot pay the mortgage. Banks require 1.25x+."),
        ("Equity Multiple", "Total money returned divided by money invested. A 2.0x means you doubled your money."),
        ("Hold Period", "How many years you plan to own the property before selling (typically 5-10 years)."),
    ]
    for term, definition in terms:
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(*PDF.NAVY)
        pdf.cell(0, 5.5, term, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9.5)
        pdf.set_text_color(*PDF.DARK)
        pdf.multi_cell(0, 5, definition)
        pdf.ln(2)

    # ===== SECTION 2: OVERVIEW =====
    pdf.add_page()
    pdf.section_title("2. The Three AI/ML Systems  -- Overview")

    pdf.body_text(
        "The tool integrates three distinct AI/ML technologies. Crucially, these are not decorative sidebars  -- "
        "each one feeds data directly into the core financial model, changing the numbers that drive the "
        "investment decision."
    )

    pdf.table(
        ["System", "Technology", "Input", "Output", "Integration"],
        [
            ["Rent Predictor", "Polynomial\nRegression", "FRED CPI\nShelter data", "Year-by-year\ngrowth rates", "Feeds into\npro forma"],
            ["ML Valuation", "Gradient\nBoosting", "13 property\nfeatures", "Fair value\nassessment", "Feeds into\nrecommendation"],
            ["Lease NLP", "Claude LLM\n+ PyPDF2", "Lease PDF\ndocuments", "Extracted terms\n+ risk flags", "Validates\nassumptions"],
        ],
        col_widths=[30, 28, 30, 32, 30],
    )

    pdf.callout_box("KEY INSIGHT: All three systems feed INTO the financial model. They change the cash flows, validate the assumptions, and influence the final BUY/HOLD/PASS recommendation.")

    pdf.body_text(
        "Without these systems enabled, the tool produces a standard DCF (Discounted Cash Flow) analysis with flat "
        "growth assumptions and an IRR-only recommendation. With them enabled, the analysis becomes multi-signal: "
        "the recommendation considers IRR strength, ML valuation, lease risk, rent outlook, and debt coverage "
        "simultaneously."
    )

    # ===== SECTION 3: RENT PREDICTOR =====
    pdf.add_page()
    pdf.section_title("3. System 1: Predictive Rent Growth Model")

    pdf.sub_title("The Problem It Solves")
    pdf.body_text(
        "When an analyst builds a financial model, they must assume how fast rents will grow each year. Most analysts "
        "type '3% per year' and leave it flat for the entire hold period. But rents do not grow at a constant rate  -- "
        "they spike during housing shortages, slow during recessions, and fluctuate with inflation. A flat assumption "
        "ignores all of this."
    )

    pdf.sub_title("How It Works")

    pdf.sub_sub_title("Step 1: Data Collection")
    pdf.body_text(
        "The tool calls the Federal Reserve Economic Data (FRED) API and downloads up to 5 years of monthly CPI "
        "Shelter data. CPI Shelter is a component of the Consumer Price Index specifically tracking housing costs. "
        "It is published monthly by the Bureau of Labor Statistics and is the most widely used measure of rent "
        "inflation in the United States."
    )

    pdf.sub_sub_title("Step 2: Data Transformation")
    pdf.body_text(
        "The raw monthly index values are converted into year-over-year growth rates. For example, if the CPI Shelter "
        "index was 300.0 in January 2023 and 315.0 in January 2024, that represents a 5.0% annual growth rate. This "
        "produces approximately 15-20 annual growth observations as training data."
    )

    pdf.sub_sub_title("Step 3: Model Training")
    pdf.body_text(
        "We fit a Polynomial Regression model (degree 2) from scikit-learn to these historical growth rates. "
        "In plain language: the model draws a curve through the historical data points and extends that curve "
        "forward in time."
    )
    pdf.body_text(
        "Why degree 2 (quadratic) rather than degree 1 (linear)? A linear model assumes growth changes at a constant "
        "pace. A quadratic curve can capture acceleration or deceleration. For instance, if rent growth has been "
        "slowing from 8% to 5% to 3%, a linear model might predict 1% next year, while the quadratic model recognizes "
        "the deceleration is itself slowing and might predict 2.5%."
    )

    pdf.sub_sub_title("Step 4: Prediction")
    pdf.body_text(
        "The model forecasts a specific growth rate for each year of the hold period. Instead of a flat '3% every "
        "year,' the output might be:"
    )
    pdf.table(
        ["", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "Year 6", "Year 7"],
        [["Growth Rate", "3.2%", "3.5%", "3.1%", "2.8%", "3.0%", "2.9%", "3.3%"]],
        col_widths=[25, 23.6, 23.6, 23.6, 23.6, 23.6, 23.6, 23.6],
    )
    pdf.body_text(
        "All predictions are clamped between -2% and +8% to prevent the model from producing absurd forecasts in "
        "extreme market conditions."
    )

    pdf.sub_title("How It Integrates Into the Financial Model")
    pdf.callout_box("CRITICAL: The predicted rates are injected BEFORE the pro forma is built. They replace the flat growth assumption.")

    pdf.body_text(
        "The pipeline works as follows:"
    )
    pdf.code_block(
        "FRED CPI Shelter data (5 years monthly)\n"
        "    -> Convert to annual growth rates (~15-20 points)\n"
        "    -> Train PolynomialRegression(degree=2)\n"
        "    -> Predict growth rates for each hold year\n"
        "    -> Inject into deal.yearly_revenue_growth = [3.2, 3.5, 3.1, ...]\n"
        "    -> build_pro_forma() reads these year-by-year\n"
        "    -> Year 1 rent grows 3.2%, Year 2 grows 3.5%, etc.\n"
        "    -> NOI, cash flows, and IRR all reflect variable growth"
    )
    pdf.body_text(
        "This means the entire 10-year projection  -- every revenue line, every NOI figure, every cash flow, and "
        "the final IRR  -- is shaped by data-driven growth predictions rather than a static guess."
    )

    pdf.sub_title("Limitations")
    pdf.bullet("CPI Shelter is a national index. It does not capture city-level rent dynamics (Austin vs. San Francisco).")
    pdf.bullet("Polynomial extrapolation can diverge over long horizons. The -2% to +8% clamp mitigates this.")
    pdf.bullet("Past inflation patterns may not predict future shocks (pandemics, policy changes).")
    pdf.bullet("These limitations are disclosed in the output.")

    # ===== SECTION 4: ML VALUATION =====
    pdf.add_page()
    pdf.section_title("4. System 2: ML Property Valuation")

    pdf.sub_title("The Problem It Solves")
    pdf.body_text(
        "When buying a property for $12.5 million, a natural question is: 'Is that price fair?' You could look at "
        "comparable sales (what similar buildings sold for recently), but reliable comp data for commercial real "
        "estate is expensive and often unavailable through free APIs. We use machine learning to generate an "
        "independent estimate of what the property should be worth."
    )

    pdf.sub_title("How It Works")

    pdf.sub_sub_title("Step 1: Synthetic Training Data Generation")
    pdf.body_text(
        "There is no free database of real commercial property transactions. So we generate 500 synthetic (simulated) "
        "properties  -- but we calibrate them to real macroeconomic data pulled from government APIs:"
    )
    pdf.bullet("10-Year Treasury rate from FRED (e.g., 4.25%)  -- anchors cap rates and property yields")
    pdf.bullet("Median household income from Census Bureau ACS (e.g., $75,000 for Austin)  -- anchors rent levels")
    pdf.bullet("National unemployment rate from FRED/BLS (e.g., 3.8%)  -- reflects economic conditions")
    pdf.bullet("30-year mortgage rate from FRED  -- affects financing and property values")

    pdf.body_text(
        "Each synthetic property has realistic, correlated features. Rents are derived from real median income "
        "(30-35% of income typically goes to rent). Cap rates are derived from the real Treasury rate plus a "
        "property-type spread. Values are derived from NOI and cap rate. While no individual property is real, "
        "the distribution is realistic for the current market environment."
    )

    pdf.sub_sub_title("Step 2: Feature Engineering (13 Features)")
    pdf.table(
        ["Feature", "Description", "Example"],
        [
            ["total_units", "Number of apartments/spaces", "50"],
            ["sf_per_unit", "Square feet per unit", "900"],
            ["year_built", "Construction year", "1995"],
            ["occupancy", "Percentage occupied", "0.92"],
            ["in_place_rent", "Current rent per unit/month", "$1,300"],
            ["market_rent", "Market rent per unit/month", "$1,450"],
            ["noi_per_unit", "Net operating income per unit", "$8,200"],
            ["property_class", "Class A=3, B=2, C=1", "2"],
            ["market_cap_rate", "Derived from Treasury + spread", "5.75%"],
            ["median_income", "City/state median income", "$75,000"],
            ["population_millions", "Metro population", "2.3"],
            ["unemployment_rate", "National unemployment", "3.8%"],
            ["mortgage_rate", "30-year mortgage rate", "6.50%"],
        ],
        col_widths=[38, 85, 27],
    )
    pdf.body_text("The target variable (what the model predicts) is value per unit in dollars.")

    pdf.sub_sub_title("Step 3: Model Architecture")
    pdf.body_text(
        "We use a Gradient Boosting Regressor from scikit-learn with the following hyperparameters:"
    )
    pdf.code_block(
        "GradientBoostingRegressor(\n"
        "    n_estimators=200,     # 200 sequential decision trees\n"
        "    max_depth=4,          # each tree has max 4 levels\n"
        "    learning_rate=0.1,    # conservative step size\n"
        "    min_samples_split=10, # prevents overfitting small groups\n"
        "    min_samples_leaf=5,   # each leaf needs 5+ samples\n"
        "    random_state=42       # reproducible results\n"
        ")"
    )
    pdf.body_text(
        "Gradient Boosting is an ensemble method: it builds hundreds of small decision trees sequentially, where "
        "each tree corrects the errors of the previous trees. It is one of the most accurate algorithms for "
        "tabular (spreadsheet-like) data and is widely used in industry."
    )

    pdf.sub_sub_title("Step 4: Honest Train/Test Split")
    pdf.callout_box("ACADEMIC HONESTY: We use an 80/20 train/test split and report held-out test metrics only.", color=(0, 100, 0))

    pdf.body_text(
        "The 500 synthetic properties are split: 400 for training, 100 held out for testing. The model never sees "
        "the test set during training. We report four metrics:"
    )
    pdf.table(
        ["Metric", "What It Measures", "Why It Matters"],
        [
            ["Train R-squared", "Fit on training data", "Sanity check (should be high)"],
            ["Test R-squared", "Fit on unseen data", "True predictive accuracy"],
            ["Test MAE", "Average dollar error", "How far off in dollars"],
            ["Test MAPE", "Average percentage error", "How far off in percent"],
        ],
        col_widths=[35, 60, 55],
    )
    pdf.body_text(
        "Many student projects report training R-squared of 0.99, which is meaningless because the model is tested "
        "on data it already memorized. Our test metrics provide an honest assessment of model capability."
    )

    pdf.sub_sub_title("Step 5: Prediction and Assessment")
    pdf.body_text(
        "Your actual deal's features are fed into the trained model. It outputs a predicted value per unit. "
        "The tool compares this to your purchase price:"
    )
    pdf.bullet("If you are paying significantly less than predicted: UNDERVALUED (potential bargain)")
    pdf.bullet("If you are paying significantly more than predicted: OVERVALUED (potential overpayment)")
    pdf.bullet("If price is within the model's error margin: FAIR VALUE")
    pdf.body_text(
        "The assessment threshold is dynamic  -- it uses the model's own test MAPE as the confidence interval. "
        "If the MAPE is 12%, then prices within +/-12% of predicted are considered FAIR VALUE."
    )

    pdf.sub_title("How It Integrates Into the Financial Model")
    pdf.body_text(
        "The ML assessment feeds into the multi-signal recommendation engine. If the model says UNDERVALUED, "
        "it adds +1 to the composite recommendation score. If OVERVALUED, it subtracts -1. This score combines "
        "with IRR, DSCR, lease analysis, and rent forecast signals."
    )

    # ===== SECTION 5: LEASE NLP =====
    pdf.add_page()
    pdf.section_title("5. System 3: Lease NLP Analysis")

    pdf.sub_title("The Problem It Solves")
    pdf.body_text(
        "Commercial leases are 30-80 page legal documents. An analyst must manually read each lease and extract "
        "key terms: What is the rent? How long is the lease? Does rent increase annually? Are there unusual clauses "
        "(early termination rights, landlord repair obligations) that create risk? For a 50-unit building, this "
        "process takes days of manual work."
    )

    pdf.sub_title("How It Works")

    pdf.sub_sub_title("Step 1: PDF Text Extraction (PyPDF2)")
    pdf.body_text(
        "The user uploads one or more lease PDF files through the web interface. PyPDF2, a Python library, reads "
        "each PDF page and converts it to raw text. This handles standard text-based PDFs. Scanned documents "
        "(images of text) may not extract properly without OCR."
    )

    pdf.sub_sub_title("Step 2: NLP Extraction via Claude API")
    pdf.body_text(
        "The extracted text is sent to Claude Sonnet (Anthropic's large language model) via API with a structured "
        "prompt. The prompt instructs Claude to extract 18 specific fields from the lease and return them as "
        "structured JSON:"
    )
    pdf.table(
        ["Field", "Description"],
        [
            ["tenant_name", "Name of the tenant"],
            ["landlord_name", "Name of the landlord/owner"],
            ["lease_type", "NNN (Triple Net), Gross, Modified Gross"],
            ["monthly_rent / annual_rent", "Rent amounts"],
            ["rent_per_sf", "Rent per square foot"],
            ["lease_term_months", "Duration of the lease"],
            ["lease_start_date / end_date", "Start and expiration dates"],
            ["escalation_clause", "Annual rent increase mechanism"],
            ["annual_escalation_pct", "Percentage increase per year"],
            ["renewal_options", "Tenant's right to extend"],
            ["security_deposit", "Deposit held by landlord"],
            ["ti_allowance", "Tenant improvement budget"],
            ["cam_charges", "Common area maintenance costs"],
            ["permitted_use", "What the space can be used for"],
            ["key_clauses", "Notable contract provisions"],
            ["risk_flags", "Concerns for the investor"],
            ["summary", "Executive summary of the lease"],
        ],
        col_widths=[55, 95],
    )

    pdf.sub_sub_title("Step 3: Regex Fallback")
    pdf.body_text(
        "If no Anthropic API key is configured, the system falls back to pattern-matching using regular expressions. "
        "It detects lease type keywords (NNN, Gross), extracts dollar amounts, identifies dates, and flags common "
        "clause types. Less accurate than Claude but functional without the API."
    )

    pdf.sub_sub_title("Step 4: Multi-Tenant Aggregation")
    pdf.body_text(
        "For buildings with multiple tenants, users can upload multiple PDFs. The tool analyzes each lease "
        "individually, then aggregates: total portfolio monthly rent, average escalation rate, combined risk "
        "flag count across all leases."
    )

    pdf.sub_title("How It Integrates Into the Financial Model")
    pdf.body_text("The lease analysis integrates at three specific points:")

    pdf.sub_sub_title("Integration Point 1: Rent Validation")
    pdf.body_text(
        "The tool compares the rent extracted from the lease to the rent the user typed into the form. If the "
        "lease says $1,500/mo but the user entered $1,300/mo as in-place rent, it flags a 15% discrepancy. This "
        "catches input errors and unrealistic assumptions before they flow through the entire model."
    )

    pdf.sub_sub_title("Integration Point 2: Growth Rate Validation")
    pdf.body_text(
        "If the lease specifies a 2% annual escalation but the user assumed 3% revenue growth, the tool flags "
        "the mismatch. The pro forma may be overstating future income relative to contractual obligations."
    )

    pdf.sub_sub_title("Integration Point 3: Risk Scoring")
    pdf.body_text(
        "The number of risk flags feeds into the recommendation engine. If Claude identifies 3+ risk flags "
        "(tenant early termination rights, landlord responsible for all repairs, short remaining term, "
        "below-market escalation), the recommendation score is penalized by -1."
    )

    # ===== SECTION 6: RECOMMENDATION ENGINE =====
    pdf.add_page()
    pdf.section_title("6. The Integrated Recommendation Engine")

    pdf.body_text(
        "Without AI/ML features enabled, the tool makes a simple recommendation based on IRR alone: "
        "above 12% = BUY, 8-12% = HOLD, below 8% = PASS. With AI/ML enabled, the recommendation becomes "
        "a composite score from five independent signals:"
    )

    pdf.table(
        ["Signal Source", "BUY (+1 or +2)", "Neutral (0)", "PASS (-1 or -2)"],
        [
            ["IRR", ">= 12% (+1)\n>= 15% (+2)", "8-12%", "< 8% (-2)"],
            ["DSCR", ">= 1.25x", "1.0-1.25x (-1)", "< 1.0x (-2)"],
            ["ML Valuation", "UNDERVALUED\n(+1)", "FAIR VALUE", "OVERVALUED\n(-1)"],
            ["Lease Analysis", "0-2 flags", " ", "3+ flags (-1)\nRent mismatch (-1)"],
            ["Rent Forecast", "> 4% avg (+1)", "1-4% avg", "< 1% avg (-1)"],
        ],
        col_widths=[35, 50, 42, 50],
    )

    pdf.sub_title("Scoring")
    pdf.table(
        ["Composite Score", "Recommendation"],
        [
            [">= 2", "STRONG BUY"],
            [">= 1", "BUY"],
            [">= 0", "HOLD / CONDITIONAL"],
            ["< 0", "PASS"],
        ],
        col_widths=[60, 90],
    )

    pdf.body_text(
        "The recommendation banner on the results page displays the final recommendation along with each "
        "signal's contribution, allowing the user to understand exactly why the tool recommends what it does."
    )

    pdf.callout_box(
        "EXAMPLE: IRR = 13.5% (BUY, +1) + DSCR = 1.35x (OK, 0) + ML says UNDERVALUED (+1) "
        "+ 1 lease flag (OK, 0) + 3.2% rent growth (OK, 0) = Score +2 = STRONG BUY"
    )

    # ===== SECTION 7: COMPLETE PIPELINE =====
    pdf.add_page()
    pdf.section_title("7. Complete Pipeline")

    pdf.body_text("The complete analysis runs in this exact order:")

    steps = [
        ("1. User Input", "Deal details entered via web form: price, units, rent, occupancy, address, financing terms, tax rate, expense overrides."),
        ("2. Market Data Fetch", "FRED API (Treasury rates, CPI, unemployment, housing), Census Bureau API (city-level demographics, income, rent), BLS API (employment data)."),
        ("3. Rent Predictor", "Trains on FRED CPI Shelter data. Predicts year-by-year growth rates. Injects rates into deal.yearly_revenue_growth BEFORE the pro forma."),
        ("4. Pro Forma Build", "10-year DCF projection using variable growth rates. Full tax analysis: depreciation, after-tax cash flows, after-tax IRR."),
        ("5. ML Valuation", "Trains GradientBoosting on 500 synthetic properties calibrated to real macro data. Predicts fair value. Classifies as UNDERVALUED / FAIR / OVERVALUED."),
        ("6. Lease NLP", "Extracts text from uploaded PDFs. Claude API extracts structured terms. Compares to user inputs. Flags discrepancies and risks."),
        ("7. Sensitivity Analysis", "4 sensitivity tables: exit cap rate, interest rate, rent growth, purchase price. Each varies the input and rebuilds the entire pro forma."),
        ("8. Recommendation", "Composite score from IRR + DSCR + ML + Lease + Rent signals. Outputs BUY/HOLD/PASS."),
        ("9. Deliverables", "Excel workbook (11 tabs) and Word investment memo generated with all results."),
    ]

    for title, desc in steps:
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(*PDF.NAVY)
        pdf.cell(0, 6, title, new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9.5)
        pdf.set_text_color(*PDF.DARK)
        pdf.multi_cell(0, 5, desc)
        pdf.ln(3)

    # ===== SECTION 8: TECH STACK =====
    pdf.add_page()
    pdf.section_title("8. Technology Stack")

    pdf.table(
        ["Component", "Technology", "Purpose"],
        [
            ["Web Framework", "Flask (Python)", "HTTP server, routing, form handling"],
            ["ML Models", "scikit-learn", "Polynomial Regression, GradientBoosting"],
            ["NLP / AI", "Anthropic Claude API", "Lease document understanding"],
            ["PDF Extraction", "PyPDF2", "Convert lease PDFs to text"],
            ["Data Manipulation", "pandas, NumPy", "Feature engineering, data processing"],
            ["Financial Math", "numpy-financial", "IRR, mortgage amortization"],
            ["Excel Generation", "openpyxl", "Multi-tab workbook with charts"],
            ["Word Generation", "python-docx", "Investment memo document"],
            ["API Integration", "requests", "FRED, Census, BLS API calls"],
            ["Frontend", "Bootstrap 5, Chart.js", "Responsive UI and charts"],
        ],
        col_widths=[38, 55, 57],
    )

    # ===== SECTION 9: DATA SOURCES =====
    pdf.section_title("9. Data Sources and API Integrations")

    pdf.table(
        ["Source", "API", "Data Retrieved"],
        [
            ["Federal Reserve (FRED)", "api.stlouisfed.org", "Treasury rates, CPI, mortgage rates,\nunemployment, housing starts"],
            ["Census Bureau", "api.census.gov", "City-level population, median income,\nmedian rent, vacancy rates (ACS 5-Year)"],
            ["Bureau of Labor\nStatistics", "api.bls.gov", "Employment data, wages,\nCPI Urban consumers"],
        ],
        col_widths=[40, 45, 65],
    )

    pdf.body_text(
        "Census data uses FIPS place codes for city-level queries on 40+ major US cities, with automatic "
        "fallback to state-level data when city codes are unavailable."
    )

    # ===== SECTION 10: HONESTY =====
    pdf.add_page()
    pdf.section_title("10. Academic Honesty and Limitations")

    pdf.sub_title("What We Disclose")
    pdf.bullet("The ML model reports TEST set metrics (80/20 split), not inflated training scores.")
    pdf.bullet("Synthetic training data is explicitly labeled  -- we never claim it is real transaction data.")
    pdf.bullet("Comparable sales are marked 'SYNTHETIC  -- NOT real transaction records.'")
    pdf.bullet("National CPI data limitations are disclosed when used for local predictions.")
    pdf.bullet("Every output states: 'This is NOT a substitute for a professional appraisal.'")
    pdf.bullet("Every data source is attributed  -- FRED, Census Bureau ACS, Bureau of Labor Statistics.")

    pdf.sub_title("Known Limitations")
    pdf.bullet("The ML model trains on synthetic data. While calibrated to real macro indicators, the R-squared reflects fit on synthetic distributions, not real transaction predictive accuracy.")
    pdf.bullet("CPI Shelter is a national index. City-level rent dynamics are not captured in the rent prediction model.")
    pdf.bullet("Lease NLP depends on PDF text quality. Scanned documents without OCR will not extract properly.")
    pdf.bullet("Census ACS data is the 5-year estimate (2018-2022). It lags current conditions by several years.")
    pdf.bullet("Polynomial extrapolation can diverge over long horizons. The -2% to +8% clamp is a safeguard, not a solution.")
    pdf.bullet("The recommendation engine uses fixed thresholds. It does not adapt to market cycles or investor risk profiles.")

    pdf.ln(6)
    pdf.callout_box(
        "DISCLAIMER: This tool is for educational and analytical purposes. All AI/ML outputs should supplement, "
        "not replace, professional judgment, independent appraisals, and proper due diligence.",
        color=(150, 0, 0),
    )

    pdf.ln(10)
    pdf.set_font("Helvetica", "I", 9)
    pdf.set_text_color(*PDF.GRAY)
    pdf.cell(0, 8, f"Document generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", align="C")

    # Save
    output_path = r"C:\ai project folder\re-underwriting-tool-clone\outputs\AI_ML_Integration_Document.pdf"
    pdf.output(output_path)
    print(f"PDF saved to: {output_path}")
    return output_path


if __name__ == "__main__":
    build_pdf()
