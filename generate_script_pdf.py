"""Generate the class presentation script as a PDF."""
from fpdf import FPDF
import os


class PDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(27, 42, 74)
        self.cell(0, 8, "RE Underwriting Tool - Presentation Script", align="R")
        self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")

    def section_title(self, title, timing=""):
        self.set_font("Helvetica", "B", 14)
        self.set_text_color(27, 42, 74)
        self.cell(0, 10, title, new_x="LMARGIN", new_y="NEXT")
        if timing:
            self.set_font("Helvetica", "I", 9)
            self.set_text_color(200, 169, 81)
            self.cell(0, 5, timing, new_x="LMARGIN", new_y="NEXT")
        self.set_draw_color(200, 169, 81)
        self.set_line_width(0.5)
        self.line(10, self.get_y(), 200, self.get_y())
        self.ln(4)

    def say(self, text):
        self.set_font("Helvetica", "", 10.5)
        self.set_text_color(50, 50, 50)
        self.multi_cell(0, 6, text)
        self.ln(2)

    def action(self, text):
        self.set_font("Helvetica", "BI", 10)
        self.set_text_color(200, 169, 81)
        self.cell(0, 6, text, new_x="LMARGIN", new_y="NEXT")
        self.ln(1)

    def bullet(self, text):
        self.set_font("Helvetica", "", 10.5)
        self.set_text_color(50, 50, 50)
        self.cell(8, 6, "-")
        self.multi_cell(0, 6, text)
        self.ln(1)

    def sub_heading(self, text):
        self.set_font("Helvetica", "B", 11)
        self.set_text_color(27, 42, 74)
        self.cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")


pdf = PDF()
pdf.alias_nb_pages()
pdf.set_auto_page_break(auto=True, margin=20)

# ---- COVER PAGE ----
pdf.add_page()
pdf.ln(50)
pdf.set_font("Helvetica", "B", 32)
pdf.set_text_color(27, 42, 74)
pdf.cell(0, 15, "RE Underwriting Tool", align="C", new_x="LMARGIN", new_y="NEXT")
pdf.ln(5)
pdf.set_font("Helvetica", "", 16)
pdf.set_text_color(200, 169, 81)
pdf.cell(0, 10, "Class Presentation Script", align="C", new_x="LMARGIN", new_y="NEXT")
pdf.ln(10)
pdf.set_font("Helvetica", "", 12)
pdf.set_text_color(100, 100, 100)
pdf.cell(0, 8, "Institutional-Grade Real Estate Investment Analysis", align="C", new_x="LMARGIN", new_y="NEXT")
pdf.ln(30)
pdf.set_font("Helvetica", "I", 10)
pdf.set_text_color(150, 150, 150)
pdf.cell(0, 8, "Total Presentation Time: ~10-12 minutes", align="C", new_x="LMARGIN", new_y="NEXT")
pdf.cell(0, 8, "Demo URL: http://localhost:5001", align="C", new_x="LMARGIN", new_y="NEXT")

# ---- 1. OPENING ----
pdf.add_page()
pdf.section_title("1. OPENING", "~1 minute")
pdf.say(
    "\"Today I'm going to walk you through a real estate investment underwriting tool I built. "
    "In practice, when institutional investors - firms like Blackstone, Starwood, or any REIB shop - "
    "evaluate an acquisition, they go through a rigorous underwriting process. They build a pro forma, "
    "model debt, calculate returns, run sensitivities, and produce deliverables for their investment "
    "committee. This tool automates that entire workflow.\""
)

# ---- 2. THE PROBLEM ----
pdf.section_title("2. THE PROBLEM IT SOLVES", "~1 minute")
pdf.say(
    "\"Typically, an analyst would spend 4-6 hours building an Excel model from scratch for each deal - "
    "pulling comps, building a 10-year cash flow projection, formatting an investment memo. "
    "This tool takes a simple deal sheet input and, within seconds, produces:\""
)
pdf.bullet("A full 10-year pro forma with value-add rent ramping")
pdf.bullet("Return metrics - levered IRR, equity multiple, DSCR, yield on cost")
pdf.bullet("An 8-tab Excel workbook with charts and sensitivity analysis")
pdf.bullet("A 10+ page Word investment memo with a BUY or PASS recommendation")
pdf.ln(2)
pdf.say("\"Let me show you how it works.\"")

# ---- 3. DEMO INPUT ----
pdf.add_page()
pdf.section_title("3. DEMO - INPUT", "~2 minutes")
pdf.say(
    "\"Here's the landing page. Notice the form is designed like a real deal sheet - "
    "the same format you'd see in a broker's offering memorandum.\""
)
pdf.action("Click \"Load Demo Values\"")
pdf.say("\"I've pre-loaded a sample deal. Let's walk through it:\"")
pdf.bullet("Multifamily Class B - a 50-unit apartment complex in Austin, Texas")
pdf.bullet("$12.5 million purchase price - that's $250,000 per unit")
pdf.bullet("Current NOI is $650,000, giving us a 5.2% going-in cap rate")
pdf.bullet("In-place rents are $1,300/month but market rents are $1,450 - $150/unit value-add opportunity")
pdf.bullet("92% occupancy - room to improve")
pdf.bullet("$200K deferred maintenance and $500K for unit upgrades - our value-add capital")
pdf.bullet("7-year hold period")
pdf.ln(2)
pdf.say(
    "\"Notice I haven't entered any debt terms, growth rates, or expense details. The tool uses "
    "institutional defaults - 65% LTV, 6.75% rate, 30-year amortization, 3% growth - but you can "
    "override any of them in the advanced panel."
)
pdf.say(
    "The key insight here is that I only entered the Current NOI. The tool backs into the expense "
    "structure automatically. In practice, you'd verify this against the actual T-12 operating "
    "statement, but for quick screening, this gets you 90% of the way there.\""
)
pdf.action("Click \"Run Underwriting Analysis\"")

# ---- 4. PROCESSING ----
pdf.add_page()
pdf.section_title("4. DEMO - PROCESSING", "~30 seconds")
pdf.say("\"The tool is now doing several things in the background:\"")
pdf.bullet("Parsing the address to identify the market - Austin, TX")
pdf.bullet("Scraping the web for comparable sales, market cap rates, and demographic data")
pdf.bullet("Building a 10-year projection with the value-add rent ramp")
pdf.bullet("Calculating all return metrics")
pdf.bullet("Generating the Excel workbook and Word memo")

# ---- 5. RESULTS ----
pdf.ln(5)
pdf.section_title("5. DEMO - RESULTS DASHBOARD", "~3 minutes")

pdf.sub_heading("Recommendation Banner:")
pdf.say(
    "\"The tool gives a BUY or PASS recommendation based on a 12% levered IRR hurdle rate - "
    "the industry standard for value-add multifamily.\""
)

pdf.sub_heading("Return Metrics:")
pdf.say("\"Four key metrics every investment committee looks at:\"")
pdf.bullet("Levered IRR - our return on equity accounting for the time value of money")
pdf.bullet("Equity Multiple - how many times we get our equity back. 1.6x means invest $5.4M, get back $8.8M")
pdf.bullet("Cash-on-Cash - our annual cash yield on equity in Year 1")
pdf.bullet("DSCR - Debt Service Coverage Ratio. Lenders want at least 1.25x. Below that, no loan")

pdf.add_page()
pdf.sub_heading("Charts:")
pdf.say(
    "\"The NOI chart shows the growth trajectory. Notice the steeper growth in Years 1-3 - "
    "that's the rent ramp from $1,300 to $1,450 as we renovate units. After stabilization, "
    "it grows at 3% annually.\""
)

pdf.sub_heading("Pro Forma Table:")
pdf.say(
    "\"This is the core of any underwriting - a 10-year annual projection. You can see the rent "
    "per unit climbing, occupancy improving from 92% to 95%, and the resulting NOI and cash flow "
    "growth. This is exactly what you'd present to an investment committee.\""
)

pdf.sub_heading("Sources & Uses:")
pdf.say(
    "\"How we're capitalizing the deal - $8.1M of debt, $5.4M of equity, covering the purchase "
    "price plus closing costs plus our $700K capex budget.\""
)

pdf.sub_heading("Exit Summary:")
pdf.say(
    "\"At Year 7, we project selling at our exit cap rate. The forward NOI divided by the exit "
    "cap gives us the sale price, minus costs and loan payoff, equals our net proceeds.\""
)

# ---- 6. DELIVERABLES ----
pdf.add_page()
pdf.section_title("6. DEMO - DELIVERABLES", "~2 minutes")
pdf.action("Click \"Download Excel Workbook\"")
pdf.say("\"This is an 8-tab institutional-quality workbook:\"")
pdf.bullet("Summary - snapshot with a NOI chart")
pdf.bullet("Sources & Uses - capital stack breakdown")
pdf.bullet("Pro Forma - the full 10-year with every line item")
pdf.bullet("Loan - amortization schedule with balance chart")
pdf.bullet("Comps - comparable sales from market research")
pdf.bullet("Sensitivity - IRR impact of exit cap changes. Green = above hurdle, Red = below")
pdf.bullet("Market - demographics and rent trends")
pdf.bullet("Assumptions - every input documented")
pdf.ln(2)
pdf.say("\"This is the same format you'd see from a CBRE or JLL investment sales team.\"")

pdf.action("Click \"Download Investment Memo\"")
pdf.say("\"And this is a 10-page Word document - the memo you'd present to a committee:\"")
pdf.bullet("Cover page, table of contents")
pdf.bullet("Executive summary with a BUY/PASS recommendation")
pdf.bullet("Property overview, market analysis")
pdf.bullet("Comp analysis with narrative")
pdf.bullet("Investment thesis highlighting the value-add strategy")
pdf.bullet("Financial summary with the NOI trajectory")
pdf.bullet("Risk assessment matrix")
pdf.bullet("Capital structure, due diligence checklist")

# ---- 7. TECHNICAL ----
pdf.add_page()
pdf.section_title("7. TECHNICAL ARCHITECTURE", "~1 minute (optional)")
pdf.say("\"For those interested in the tech stack:\"")
pdf.bullet("Backend: Python with Flask - handles the financial modeling engine")
pdf.bullet("Financial Model: Custom 10-year DCF with numpy-financial for IRR calculations")
pdf.bullet("Excel: openpyxl generates the workbook with embedded charts")
pdf.bullet("Word: python-docx builds the formatted memo")
pdf.bullet("Market Research: Web scraping with BeautifulSoup for real-time comp and demographic data")
pdf.bullet("Frontend: Bootstrap 5 with Chart.js for the dashboard visualizations")
pdf.ln(2)
pdf.say(
    "\"The entire app is about 3,000 lines of code and runs on any machine with Python installed.\""
)

# ---- 8. CLOSING ----
pdf.ln(5)
pdf.section_title("8. CLOSING", "~30 seconds")
pdf.say(
    "\"To summarize - this tool takes a deal from a one-page deal sheet to a full institutional "
    "underwriting package in under 30 seconds. In practice, this is the kind of tool that would "
    "sit on an acquisitions team's desktop and let them screen 10 deals in the time it used to "
    "take to underwrite one. Thank you.\""
)

# ---- SAVE ----
desktop = r"C:\Users\Peter M John\OneDrive\Desktop"
output_path = os.path.join(desktop, "RE_Underwriting_Presentation_Script.pdf")
pdf.output(output_path)
print(f"PDF saved: {output_path}")
