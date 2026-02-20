# RE Underwriting Intelligence Platform

## The Problem

Commercial real estate underwriting is done almost entirely in Excel — static, manual,
error-prone, and disconnected from live market data. Analysts spend hours pulling data
from FRED, CoStar, and broker reports by hand, then hard-code assumptions that are
stale by the time a deal closes. The resulting models are opaque, hard to audit, and
produce no signal about what they don't know.

Specifically:
- No live integration with macroeconomic data (FRED treasury rates, CPI, vacancy)
- No city-level rental market benchmarks (RentCast, HUD FMR, Zillow ZORI)
- No ML-based valuation cross-check against market conditions
- No probabilistic output — a single IRR number hides all the uncertainty
- No AI-powered deal narrative — memos are copy-paste templates
- No location risk factors (crime, walkability, flood zone)
- No lease document parsing — analysts read PDFs manually

## Success Looks Like

A deal analyst enters a property address, purchase price, unit count, and basic
assumptions into a web form. Within 60 seconds they receive:

1. A full 10-year pro forma with real-market-data-calibrated rent growth
2. ML valuation cross-check (overvalued / fair / undervalued vs. market comps)
3. Monte Carlo return distribution (not just a point IRR — a probability curve)
4. Live market data: RentCast rents, HUD Fair Market Rents, FRED macro signals,
   crime risk, Walk Score, demographic profile
5. AI-generated investment memo and market narrative grounded in actual data
6. Downloadable Excel, Word, and PDF outputs
7. A chat interface to ask questions about the deal

## How We'd Know We're Wrong

- If the ML valuation model consistently produces nonsensical premiums/discounts
  vs. broker opinions of value → retrain or replace the model
- If the AI-generated memos contain factual errors about the market data →
  tighten the grounding prompt and cross-check citations
- If users distrust the outputs because they can't audit the sources →
  add data lineage labels to every number (source + date)
- If the tool is slower than Excel for simple deals → optimize the pipeline

## Building On (Existing Foundations)

- **numpy-financial** — IRR, NPV, loan amortization calculations
- **scikit-learn** — GradientBoosting for ML valuation, PolynomialRegression for rent prediction
- **FRED API (St. Louis Fed)** — 10yr Treasury, mortgage rates, CPI Shelter, housing starts,
  rental vacancy rate, national unemployment
- **Census Bureau ACS API** — City-level population, median income, median rent,
  vacancy rate, renter occupancy %
- **BLS API** — National employment and unemployment data
- **Zillow ZORI** — City-level rent index and annual rent growth (bulk CSV)
- **RentCast API** — Live zip-level market rent stats, AVM rent estimate, real rental comps
- **HUD Fair Market Rents API** — 2026 FMR by metro area and bedroom count
- **FBI UCR Crime Data** — Static 68-city violent crime dataset (Marshall Project)
- **Walk Score API** — Walk, transit, and bike scores by address
- **Anthropic Claude API** — AI deal memo, market narrative, lease analysis, chat interface
- **Google Gemini API** — PDF lease document extraction and parsing
- **Flask + Jinja2** — Web framework (Python-native, no TypeScript overhead)
- **pandas + openpyxl + python-docx + fpdf2** — Report generation

## The Unique Part

What we are actually building on top of these foundations:

1. **Integrated underwriting pipeline** — A single analysis job that calls all data
   sources in parallel, blends them into a coherent financial model, and produces
   a multi-format output package in under 90 seconds

2. **ML-calibrated rent growth forecasting** — Polynomial regression on FRED CPI Shelter
   + Zillow ZORI, blended 60/40, producing variable year-by-year growth rates instead
   of a flat assumption

3. **Multi-signal recommendation engine** — Combines IRR signal, DSCR health, ML
   valuation assessment, lease risk flags, rent forecast signal, and Monte Carlo
   probability into a scored BUY/HOLD/PASS recommendation

4. **AI-grounded deal chat interface** — Claude with the full deal context (pro forma,
   market data, ML outputs) as system prompt, letting analysts interrogate the deal
   in natural language

5. **HUD FMR vs. in-place rent comparison** — Automated flagging of below-market
   (value-add signal) and above-market (risk signal) positioning

6. **Location risk dashboard** — Crime risk band, Walk Score, HUD metro context
   synthesized into underwriting notes

## Tech Stack

- **UI:** Flask + Jinja2 + Bootstrap 5 (dark theme)
- **Backend:** Python 3.11, background threading for async analysis jobs
- **ML/Analytics:** scikit-learn, numpy, pandas, numpy-financial
- **AI:** Anthropic Claude (claude-sonnet-4-6), Google Gemini (gemini-2.0-flash)
- **Data APIs:** FRED, Census ACS, BLS, Zillow ZORI, RentCast, HUD, Walk Score, FBI UCR
- **Report Generation:** openpyxl (Excel), python-docx (Word), fpdf2 (PDF)
- **CI/CD:** GitHub Actions → Render.com (auto-deploy on master push)
- **Testing:** pytest (51 tests), flake8 linting

## System Context

```
User (Analyst)
     │
     ▼
Flask Web App (app.py)
     │
     ├── Background Analysis Thread
     │        │
     │        ├── Market Research (parallel)
     │        │        ├── FRED API → Treasury, CPI, mortgage, vacancy
     │        │        ├── Census ACS → demographics
     │        │        ├── BLS → employment
     │        │        ├── Zillow ZORI → rent index
     │        │        ├── RentCast API → live comps, AVM
     │        │        ├── HUD API → Fair Market Rents
     │        │        ├── Walk Score API → location scores
     │        │        └── Crime CSV → risk band
     │        │
     │        ├── Rent Predictor (polynomial regression on FRED+ZORI)
     │        ├── Financial Model (10-year pro forma)
     │        ├── ML Valuation (GradientBoosting)
     │        ├── Lease Analyzer (Gemini/Claude PDF extraction)
     │        ├── Sensitivity Analysis (4 tables)
     │        ├── Monte Carlo (1000 iterations)
     │        ├── Recommendation Engine (multi-signal scoring)
     │        └── Report Generation (Excel + Word + PDF)
     │
     └── Results Page → Deal Chat Interface (Claude, grounded in deal context)
```

## Open Questions (In-Progress)

- [ ] Walk Score API key pending (need peterjohn1298.github.io domain validation)
- [ ] Deal Chat Interface — not yet built (next priority)
- [ ] AI-generated Deal Memo — partially built (Word template), needs Claude integration
- [ ] Unit mix modeling by bedroom count — not yet built
- [ ] Waterfall / LP-GP promote structure — not yet built
- [ ] Database persistence — deals currently in-memory only
- [ ] OpenFEMA flood zone integration — not yet built
