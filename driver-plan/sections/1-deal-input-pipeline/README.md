# Section 1: Deal Input & Analysis Pipeline

## What It Does
Collects all deal parameters via a multi-tab Streamlit form and runs the full underwriting
pipeline, storing results in session_state and navigating to the Results Dashboard.

## Key Files
- `streamlit_app.py` — Entry point, form, pipeline execution

## Inputs (Form Tabs)
| Tab | Fields |
|-----|--------|
| Deal Sheet | address, property type, year built, purchase price, current NOI, units, SF, in-place rent, market rent, occupancy, hold period, deferred maintenance, capex |
| Unit Mix | Studio/1BR/2BR/3BR counts, in-place rents, market rents (with live blended rent preview) |
| AI/ML Features | enable ML valuation, enable rent prediction, enable waterfall, lease PDF uploads |
| Override Defaults | LTV, interest rate, amortization, loan term, I/O period, closing/sale costs, growth rates, management fee, exit cap spread, tax rate, land value, LP/GP waterfall params, expense overrides |

## Pipeline Steps (in order)
1. Derive assumptions (`models/assumptions.py`)
2. Market research — 7 APIs (`services/market_research.py`)
3. Rent prediction — ML (`models/rent_predictor.py`) [optional]
4. Pro forma — DCF, IRR, EM, DSCR, CoC (`models/financial_model.py`)
5. ML valuation — GradientBoosting (`models/ml_valuation.py`) [optional]
6. Lease analysis — AI PDF extraction (`services/lease_analyzer.py`) [optional]
7. Sensitivity tables (`models/metrics.py`)
8. Monte Carlo — 1,000 simulations (`models/monte_carlo.py`)
9. Recommendation — signal scoring (`streamlit_app.py`)
10. Waterfall — LP/GP distribution (`models/waterfall.py`) [optional]
11. AI memo — GPT-4o (`services/ai_memo.py`) [if OPENAI_API_KEY set]
12. Base/Bull/Bear scenarios (`models/scenarios.py`)
13. Excel, Word, PDF generation (`services/excel_generator.py` etc.)

## Demo Button
Pre-fills all fields with Austin TX 100-unit Class B deal ($12.5M) including unit mix,
and enables ML + rent prediction checkboxes.

## Status: COMPLETE
