# RE Underwriting Intelligence Platform

## The Problem
Underwriting a multifamily real estate deal requires pulling data from a dozen sources,
running a 10-year DCF, stress-testing assumptions, and writing an investment memo — all
manually in Excel. This takes days, is error-prone, and produces results that are hard to
share or reproduce.

## Success Looks Like
Enter an address + deal terms → one click → institutional-quality underwriting package
available at a public URL, shareable with any stakeholder.

## Five Sections

| # | Section | Entry Point |
|---|---------|-------------|
| 1 | Deal Input & Analysis Pipeline | `streamlit_app.py` |
| 2 | Results Dashboard (11 tabs) | `pages/1_Results.py` |
| 3 | AI Deal Chat (Claude) | `pages/2_Chat.py` |
| 4 | Deal Comparison | `pages/3_Compare.py` |
| 5 | Report Export (Excel/Word/PDF) | `services/*_generator.py` |

## Tech Stack
- **UI:** Streamlit (multi-page, dark navy/gold theme, Plotly charts)
- **Finance Models:** Custom Python (IRR, DCF, Monte Carlo, waterfall, scenarios)
- **ML:** scikit-learn GradientBoosting (valuation), Ridge/Linear (rent prediction)
- **AI:** OpenAI GPT-4o (memo), Anthropic Claude claude-opus-4-6 (chat), Google Gemini (leases)
- **Data APIs:** FRED, RentCast, HUD, Walk Score, FBI, FEMA, Census/BLS
- **Reports:** openpyxl (Excel), python-docx (Word), reportlab (PDF)
- **Deployment:** Streamlit Community Cloud

## Key Design Decisions
- All models in `models/` are pure Python with zero UI dependency — testable independently
- Session state (`st.session_state["results"]`) is the data bus between pages
- `_secret()` helper resolves API keys from `st.secrets` → `os.environ` (local dev)
- Monte Carlo returns IRR values already as % (not 0–1) — do not multiply by 100
- Scenarios returns flat dicts per scenario + pre-built `comparison_table` list
- ML valuation guard: check `ml.get("assessment")` not `ml.get("available")`
