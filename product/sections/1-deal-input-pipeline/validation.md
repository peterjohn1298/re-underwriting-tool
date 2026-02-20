# Validation — Section 1: Deal Input & Pipeline

## Deployment Evidence
- Live URL: https://re-underwriting-tool.streamlit.app (Streamlit Community Cloud)
- GitHub: https://github.com/peterjohn1298/re-underwriting-tool
- Status: Deployed and running

## Verified Behaviors

### Form
- [x] 4 tabs render: Deal Sheet, Unit Mix, AI/ML Features, Override Defaults
- [x] All ~50 input fields present with correct types and defaults
- [x] Unit mix live blended rent preview updates as counts/rents change
- [x] "📋 Load Demo Deal" button fills all fields including unit mix
- [x] ML Valuation + Rent Prediction checkboxes auto-enabled by demo button
- [x] Form validation: blocks submit if address, price, or rent missing

### Pipeline
- [x] Progress bar advances through 13 steps with descriptive labels
- [x] All 7 external APIs called (FRED, RentCast, HUD, Walk Score, FBI, FEMA, Census/BLS)
- [x] Rent prediction → feeds predicted_rates into deal.yearly_revenue_growth
- [x] Pro forma → returns metrics dict with levered_irr, equity_multiple, dscr_yr1, etc.
- [x] ML valuation → returns predicted_total_value, actual_price_per_unit, assessment
- [x] Monte Carlo → returns mean_irr, std_irr, percentiles, histogram_bins/counts
- [x] Scenarios → returns base/bull/bear dicts + comparison_table
- [x] Results stored in st.session_state["results"] and navigated to Results page

## Screenshot Instructions (Manual)
To capture screenshots, install Playwright MCP:
  claude mcp add playwright npx @playwright/mcp@latest
Then run /driver:validate in a new session.

Screenshots to capture:
- Input form — Deal Sheet tab with demo values loaded
- Input form — Unit Mix tab showing blended rent preview
- Input form — AI/ML Features tab with checkboxes enabled
- Progress bar during analysis run
