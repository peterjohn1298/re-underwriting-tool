# Section 5: Report Export

## What It Does
Generates institutional-quality Excel, Word, and PDF reports from the pipeline output dict,
available as one-click downloads on the Results Dashboard.

## Key Files
- `services/excel_generator.py` — Excel workbook (openpyxl)
- `services/word_generator.py` — Word memo (python-docx)
- `services/pdf_generator.py` — PDF report (reportlab)

## Excel Workbook Contents
- Pro forma (10-year: EGI, expenses, NOI, debt service, BTCF, ATCF)
- Sensitivity tables (exit cap, interest rate, rent growth, purchase price vs IRR)
- Market data (demographics, cap rates, crime, walk score, HUD FMRs)
- ML valuation results + feature importances
- Monte Carlo percentile table + probability table
- Rent prediction year-by-year table
- Backtest results

## Word Memo Contents
- Property overview and capital structure
- Executive summary narrative
- Market analysis section
- Financial analysis with key metrics
- Risk factors
- Investment thesis
- Recommendation with signal breakdown
- AI-generated memo sections (if OpenAI key present)
- Unit mix summary

## PDF Report Contents
- One-page summary with key metrics and verdict
- Pro forma table
- Market data highlights
- Monte Carlo summary

## Download Format
Reports named: `underwriting_{job_id}.xlsx`, `investment_memo_{job_id}.docx`, `investment_report_{job_id}.pdf`
Generated paths stored in `session_state["results"]["excel_path"]` etc.

## Status: COMPLETE
