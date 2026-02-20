# Section 2: Results Dashboard

## What It Does
Displays all 11 analysis outputs across tabbed sections with interactive charts, metric cards,
a real-time what-if sidebar calculator, and navigation to Chat, Compare, and Downloads.

## Key Files
- `pages/1_Results.py` — Full results dashboard

## Tabs
| # | Tab | Key Content |
|---|-----|-------------|
| 0 | Overview | Verdict badge (STRONG BUY/BUY/HOLD/PASS), score, 5 scorecard chips, signal breakdown table, unit mix table |
| 1 | Returns | IRR (levered/unlevered/after-tax), equity multiple, CoC, DSCR, YOC, capital structure, exit/reversion metrics |
| 2 | Pro Forma | 10-year table (EGI, expenses, NOI, debt service, BTCF, ATCF), amortization schedule, sensitivity tables |
| 3 | Market Data | Demographics, FRED macro rates, RentCast comps/AVM, HUD FMRs, Walk Score, FEMA flood zone, crime risk |
| 4 | ML & Monte Carlo | ML valuation badge + metrics, feature importances, MC percentile table + probability table + IRR histogram |
| 5 | Rent Forecast | Year-by-year rent/growth/revenue table, Plotly line chart, model backtest results |
| 6 | Scenarios | Pre-built Base/Bull/Bear comparison table, side-by-side IRR + EM bar charts |
| 7 | Lease Analysis | Portfolio summary, per-lease extracted terms, risk flags, input comparison flags |
| 8 | AI Memo | 6-section expandable memo (executive summary, market, financial, risks, thesis, recommendation) |
| 9 | Waterfall | LP/GP IRR, total return, multiple; distribution tiers table |
| 10 | Downloads | Excel, Word, PDF download buttons |

## Sidebar (What-If Calculator)
Sliders for rent growth, occupancy, exit cap spread, expense growth → recalculates full
pro forma in real-time using `build_pro_forma()` and updates IRR/EM/CoC/DSCR metrics.

## Header Buttons
- **💬 AI Chat** → navigates to pages/2_Chat.py
- **← New Analysis** → navigates to streamlit_app.py
- **📌 Save for Comparison** → saves deal to session_state["saved_deals"]

## Status: COMPLETE
