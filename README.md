# RE Underwriting Intelligence Platform

**Institutional-grade multifamily real estate underwriting, powered by AI & ML.**

[![CI](https://github.com/peterjohn1298/re-underwriting-tool/actions/workflows/ci.yml/badge.svg)](https://github.com/peterjohn1298/re-underwriting-tool/actions/workflows/ci.yml)

Live demo: [re-underwriting-tool.streamlit.app](https://re-underwriting-tool.streamlit.app)

---

## Overview

This tool automates the full underwriting workflow for multifamily real estate acquisitions.
Enter an address and deal parameters → one click → complete investment package including
a 10-year pro forma, ML-based property valuation, Monte Carlo risk simulation, housing
market regime detection, AI lease analysis, and an AI-generated investment memo.

**No Bloomberg terminal. No Excel macros. No local Python environment required.**

---

## AI & ML Features

This project was built for *Mastering AI for Finance*. Every major analytical component
uses machine learning or AI:

### 1. ML Property Valuation (scikit-learn GradientBoosting)
- Trains a `GradientBoostingRegressor` on market comparable data fetched at runtime
- Predicts fair value per unit from features: location, unit mix, year built, income, cap rate
- Classifies the subject property as **UNDERVALUED / FAIR VALUE / OVERVALUED**
- Reports `premium_discount_pct` and top valuation drivers via `feature_importances_`
- Model accuracy (R², MAE, MAPE) reported in the Results dashboard

### 2. ML Rent Growth Prediction (Ridge Regression + ZORI)
- Trains on historical CPI Shelter (FRED CUSR0000SAH1) and Zillow ZORI growth rates
- Predicts year-by-year rent growth for the full hold period
- Predictions feed directly into the pro forma (replaces flat growth rate assumption)
- Backtest results (MAE, RMSE, R²) shown alongside predictions

### 3. Housing Market Regime Detection (Rule-Based ML Scoring)
- Classifies the macro housing cycle as **EXPANSION / PEAK / CONTRACTION / TROUGH**
- Uses five live FRED signals: housing starts trend, rental vacancy rate, CPI Shelter YoY,
  yield curve spread (10yr–2yr), and unemployment rate
- Each signal scored –1 / 0 / +1; composite score maps to regime
- Regime displayed on Overview and Market Data tabs with underwriting implications
- Directly addresses regime detection methodology (see: housing market regime shifts)

### 4. Monte Carlo Simulation (1,000 paths)
- Simulates IRR distribution across 1,000 randomized market scenarios
- Random variables: rent growth, vacancy, exit cap rate, expense growth
- Outputs: P10/P50/P90 IRR, mean, std dev, probability table, full histogram
- Helps quantify downside risk beyond deterministic Base/Bull/Bear scenarios

### 5. AI Lease Document Analysis (LLM — Google Gemini / OpenAI GPT-4o)
- Extracts structured terms from uploaded PDF leases: monthly rent, term, escalation, expiry
- Flags risk clauses: below-market rents, co-tenancy provisions, early termination rights
- Compares extracted rents against underwriting assumptions; flags discrepancies > 10%
- Supports single lease or portfolio of leases

### 6. AI Investment Memo (OpenAI GPT-4o)
- Generates a 6-section institutional investment memo grounded in all analysis outputs
- Sections: Executive Summary, Market Analysis, Financial Analysis, Risk Factors,
  Investment Thesis, Recommendation
- System prompt includes all computed metrics, market data, and ML outputs

### 7. AI Deal Chat (Anthropic Claude claude-opus-4-6)
- Conversational analyst grounded entirely in the deal's analysis results
- Cannot hallucinate numbers — system prompt contains all actual computed outputs
- Accessible via "💬 AI Chat" button on the Results page
- Maintains 10-turn conversation history per session

---

## Data Sources (Real APIs — No Web Scraping)

The previous version of this tool used Google web scraping, which was blocked and fell back
to synthetic data. This version uses proper API integrations throughout:

| API | Data | Key |
|-----|------|-----|
| **FRED** (Federal Reserve) | 10-yr treasury, 30-yr mortgage, CPI, unemployment, housing starts, rental vacancy | `FRED_API_KEY` |
| **RentCast** | AVM rent estimate, market comps, median rents, days on market | `RENTCAST_API_KEY` |
| **HUD** | Fair Market Rents by bedroom type (used in unit mix analysis) | `HUD_API_KEY` |
| **Walk Score** | Walkability, transit, and bike scores | `GOOGLE_API_KEY` |
| **FBI UCR** | Crime risk index vs national average | `FBI_CRIME_API_KEY` |
| **FEMA** | Flood zone classification (SFHA / non-SFHA) | — (public) |
| **Census / BLS** | Population, median income, renter %, unemployment | — (public) |
| **OpenAI** | Investment memo generation (GPT-4o) | `OPENAI_API_KEY` |
| **Anthropic** | Deal chat assistant (Claude claude-opus-4-6) | `ANTHROPIC_API_KEY` |
| **Google Gemini** | Lease PDF analysis (fallback: OpenAI) | `GOOGLE_API_KEY` |

---

## Financial Model

All calculations use established real estate finance methodology:

- **IRR**: Levered, unlevered, and after-tax (numpy-financial + custom Newton solver)
- **Equity Multiple**: Total distributions / equity invested
- **DSCR**: NOI / annual debt service (Year 1 and average over hold)
- **Cash-on-Cash**: Before-tax cash flow / equity invested
- **Yield on Cost**: Stabilized NOI / total project cost
- **DCF**: 10-year pro forma with amortizing debt, depreciation shield, capital gains tax
- **Reversion**: Exit cap applied to Year N+1 forward NOI, net of selling costs and taxes
- **LP/GP Waterfall**: Preferred return hurdle, catch-up, promote (carried interest)
- **Sensitivity**: IRR vs exit cap rate, interest rate, rent growth, purchase price

---

## Architecture

```
streamlit_app.py           ← Entry point: input form + 13-step pipeline
pages/
  1_Results.py             ← 11-tab results dashboard + what-if sidebar
  2_Chat.py                ← Claude AI deal chat
  3_Compare.py             ← Side-by-side deal comparison
models/
  assumptions.py           ← DealInputs dataclass + derive_assumptions()
  financial_model.py       ← Core DCF / pro forma engine
  metrics.py               ← IRR, EM, CoC, DSCR, sensitivity tables
  ml_valuation.py          ← GradientBoosting property valuation
  rent_predictor.py        ← Ridge regression rent growth forecast
  monte_carlo.py           ← 1,000-path IRR simulation
  regime_detector.py       ← Housing market cycle classification (FRED)
  scenarios.py             ← Base / Bull / Bear scenario engine
  waterfall.py             ← LP/GP distribution model
  unit_mix.py              ← Unit type → blended rent calculation
  backtest.py              ← Rent model backtester
services/
  market_research.py       ← Orchestrates all 7+ API calls
  lease_analyzer.py        ← LLM-based PDF lease extraction
  ai_memo.py               ← GPT-4o investment memo generator
  excel_generator.py       ← openpyxl Excel workbook
  word_generator.py        ← python-docx investment memo
  pdf_generator.py         ← reportlab PDF summary
  api_clients/
    fred_client.py         ← FRED API (macro data)
    rentcast_client.py     ← RentCast API (rent comps)
    hud_client.py          ← HUD API (fair market rents)
    walkscore_client.py    ← Walk Score API
    crime_client.py        ← FBI UCR crime data
    fema_client.py         ← FEMA flood zone service
    census_client.py       ← Census / ACS demographics
    bls_client.py          ← BLS unemployment / labor data
    zillow_client.py       ← Zillow ZORI rent index
```

---

## Installation & Local Development

```bash
# 1. Clone
git clone https://github.com/peterjohn1298/re-underwriting-tool
cd re-underwriting-tool

# 2. Install dependencies
pip install -r requirements.txt

# 3. Add API keys
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Edit .streamlit/secrets.toml and fill in your API keys

# 4. Run
streamlit run streamlit_app.py
```

### API Keys

Copy `.streamlit/secrets.toml.example` → `.streamlit/secrets.toml` and fill in:

```toml
FRED_API_KEY        = "..."   # https://fred.stlouisfed.org/docs/api/api_key.html
RENTCAST_API_KEY    = "..."   # https://rentcast.io
HUD_API_KEY         = "..."   # https://www.huduser.gov/portal/dataset/fmr-api.html
GOOGLE_API_KEY      = "..."   # https://aistudio.google.com
FBI_CRIME_API_KEY   = "..."   # https://api.data.gov
OPENAI_API_KEY      = "..."   # https://platform.openai.com
ANTHROPIC_API_KEY   = "..."   # https://console.anthropic.com
```

---

## Deployment

Deployed on **Streamlit Community Cloud** — zero infrastructure, auto-deploys on push.

1. Fork this repo
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your repo, set main file to `streamlit_app.py`
4. Add all API keys in the Secrets editor (TOML format)
5. Deploy → get a public URL instantly

---

## Testing

```bash
pytest tests/ -v
```

51 tests covering financial models, API client mocking, unit mix parsing, and scenario logic.
All backend models are pure Python with zero UI dependency — fully testable without Streamlit.

---

## AI Leverage Assessment

| Component | AI/ML Method | Model | Purpose |
|-----------|-------------|-------|---------|
| Property Valuation | Gradient Boosting | scikit-learn GBR | Predict fair value, flag over/underpricing |
| Rent Forecast | Ridge Regression | scikit-learn Ridge | Predict year-by-year rent growth from macro data |
| Regime Detection | Multi-signal scoring | Custom (FRED) | Classify housing market cycle phase |
| Risk Simulation | Monte Carlo | Custom (numpy) | Quantify IRR distribution across 1,000 scenarios |
| Lease Analysis | Large Language Model | Google Gemini / GPT-4o | Extract and risk-flag lease terms from PDFs |
| Investment Memo | Large Language Model | OpenAI GPT-4o | Generate institutional-quality narrative memo |
| Deal Chat | Large Language Model | Anthropic Claude claude-opus-4-6 | Grounded Q&A about deal results |

**Total AI/ML components: 7** across supervised learning, unsupervised signal scoring,
Monte Carlo simulation, and three distinct LLM integrations.

---

## Key Improvements Over Original Submission

| Issue (Professor Feedback) | Resolution |
|---------------------------|-----------|
| Web scraping broken (Google blocks) | Replaced with 7 proper API integrations |
| No AI/ML integration | 7 AI/ML components added (see table above) |
| No README | This document |
| No regime detection | `models/regime_detector.py` — EXPANSION/PEAK/CONTRACTION/TROUGH |
| No AI documentation | AI Leverage Assessment table above |

---

## Course Context

Built for *Mastering AI for Finance*. The tool applies AI/ML techniques from the course
to a real-world institutional finance workflow:

- **Supervised learning**: property valuation and rent forecasting
- **Unsupervised / rule-based ML**: regime detection using multi-factor scoring
- **Probabilistic simulation**: Monte Carlo IRR distribution
- **NLP / LLMs**: lease extraction, memo generation, deal Q&A
- **Financial engineering**: DCF, IRR, waterfall, scenario analysis
