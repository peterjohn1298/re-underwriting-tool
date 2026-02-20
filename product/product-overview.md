# RE Underwriting Intelligence Platform

## The Problem
Underwriting a multifamily real estate deal requires pulling data from a dozen sources
(FRED, HUD, RentCast, Census, FEMA, Walk Score, FBI Crime), running a 10-year pro forma,
stress-testing assumptions, and writing an investment memo — all manually, all in Excel.
This takes days for an analyst and is error-prone, inconsistent, and hard to share.

## Success Looks Like
Enter an address + deal terms → one click → institutional-quality underwriting package:
- 10-year pro forma with IRR, EM, DSCR, CoC
- ML property valuation (fair/over/undervalued)
- Monte Carlo simulation (1,000 paths, IRR distribution)
- ML rent growth forecast (year-by-year, feeds pro forma)
- Base/Bull/Bear scenario comparison
- Market intelligence (demographics, cap rates, crime, flood zone, walk score, HUD FMRs)
- AI lease analysis (PDF extraction, risk flags)
- AI investment memo (GPT-4o)
- AI deal chat (Claude claude-opus-4-6, grounded in results)
- Downloadable Excel, Word, and PDF reports
- Shareable via public URL (Streamlit Community Cloud)

## Building On (Existing Foundations)
- **scikit-learn GradientBoosting** — ML property valuation model
- **numpy-financial / custom IRR solver** — levered/unlevered/after-tax IRR, equity multiple
- **FRED API** — 10-year treasury, 30-yr mortgage, CPI, unemployment, housing starts
- **RentCast API** — AVM rent estimate, market comps, median rents
- **HUD API** — Fair Market Rents by bedroom type
- **Walk Score API** — Walkability, transit, bike scores
- **FBI UCR data** — Crime risk scoring vs national average
- **FEMA flood map service** — Flood zone classification
- **Census / BLS** — Demographics, income, population
- **OpenAI GPT-4o** — Investment memo generation
- **Anthropic Claude claude-opus-4-6** — Grounded deal chat assistant
- **Streamlit** — UI framework, multi-page app, Community Cloud deployment
- **Plotly** — Interactive charts (Monte Carlo histogram, rent forecast, scenario bars)
- **openpyxl / python-docx / reportlab** — Excel, Word, PDF export

## The Unique Part
A single-URL, zero-install tool that integrates 7 external APIs, 4 ML models, and 3 AI
services into one underwriting workflow — accessible to anyone with a browser. No Bloomberg
terminal, no Excel macros, no local Python environment required.

## Tech Stack
- **UI:** Streamlit (multi-page: Input → Results → Chat → Compare)
- **Backend:** Pure Python models and services (zero Flask dependency in Streamlit layer)
- **ML:** scikit-learn (GradientBoostingRegressor for valuation, Linear/Ridge for rent)
- **Finance:** Custom IRR solver, DCF, amortization, Monte Carlo, waterfall distribution
- **Data:** 7 external API clients (FRED, RentCast, HUD, Walk Score, FBI, FEMA, Census/BLS)
- **AI:** OpenAI GPT-4o (memo), Anthropic Claude (chat), Google Gemini (lease analysis fallback)
- **Deployment:** Streamlit Community Cloud (public URL, secrets management)

## Open Questions
- Walk Score API free tier: 5,000 req/day — may need caching for high traffic
- FBI crime data currently uses local CSV (UCR 2019) — consider live API integration
- ML valuation model trains on market comp data at runtime — could pre-train and cache
- PDF generator may need WeasyPrint or wkhtmltopdf on Streamlit Cloud (dependency issue)
