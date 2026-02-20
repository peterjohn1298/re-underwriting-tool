# Validation — Section 2: Results Dashboard

## Deployment Evidence
- Live URL: https://re-underwriting-tool.streamlit.app → Results page after analysis
- File: pages/1_Results.py

## Verified Behaviors

### Navigation & Layout
- [x] Guard: redirects to input form if no results in session_state
- [x] Property name and address shown in header
- [x] "💬 AI Chat" button navigates to Chat page
- [x] "← New Analysis" button navigates back to input form
- [x] "📌 Save for Comparison" saves deal to session_state["saved_deals"]
- [x] What-if sidebar recalculates IRR/EM/CoC/DSCR in real time

### Tab Contents (all verified)
- [x] Overview: verdict badge, score, 5 scorecard chips, signal table, unit mix table
- [x] Returns: all 8 return metrics, capital structure, exit/reversion details
- [x] Pro Forma: 10-year table with dollar formatting, amortization schedule
- [x] Market Data: demographics, FRED rates, RentCast, HUD FMR, Walk Score, FEMA, crime
- [x] ML & Monte Carlo: valuation badge using actual keys (assessment, predicted_total_value); MC histogram using histogram_bins/histogram_counts
- [x] Rent Forecast: table from predicted_rates + predicted_rents_per_unit + predicted_annual_revenue; Plotly line chart
- [x] Scenarios: comparison_table rendered directly; IRR + EM bar charts
- [x] Lease Analysis: portfolio summary, per-lease table, risk flags
- [x] AI Memo: 6-section expandable blocks
- [x] Waterfall: LP/GP metrics, distribution tiers table
- [x] Downloads: Excel, Word, PDF download buttons with file existence check

## Key Bug Fixes Applied
- MC keys: mean_irr / std_irr / percentiles["P10"] / histogram_bins+counts
- ML keys: assessment / predicted_total_value / actual_price_per_unit / feature_importances dict
- Rent keys: predicted_rates + predicted_rents_per_unit (not yearly_predictions)
- Scenarios: comparison_table used directly; only base/bull/bear iterated for charts

## Screenshot Instructions (Manual)
Screenshots to capture:
- Overview tab with STRONG BUY verdict badge
- Returns tab with all metric cards
- ML & Monte Carlo tab with histogram
- Scenarios tab with comparison table and bar chart
- What-if sidebar with recalculated metrics
