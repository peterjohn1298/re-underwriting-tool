# Full Implementation Guide — RE Underwriting Intelligence Platform

## Overview
A single-URL Streamlit app that runs institutional-grade multifamily underwriting
in one click. No Bloomberg terminal, no Excel macros, no local Python environment.

## Repository
- GitHub: https://github.com/peterjohn1298/re-underwriting-tool
- Live: https://re-underwriting-tool.streamlit.app
- Stack: Python 3.11+, Streamlit ≥ 1.41, Plotly ≥ 5.0, scikit-learn, openpyxl, python-docx

---

## Milestone 1: Deal Input & Analysis Pipeline (`streamlit_app.py`)

### What to build
A Streamlit form with 4 tabs (~50 fields total) that validates input, builds a
`DealInputs` dataclass, runs a 13-step pipeline, and stores all results in
`st.session_state["results"]` before navigating to the Results page.

### Key implementation notes
- Use `st.form("deal_form")` with `st.tabs()` inside
- Unit mix widgets need explicit `key=f"um_count_{ut}"` — set session state directly
  for demo button since `value=` is ignored for keyed widgets after first render
- Helper functions (`_build_recommendation`, `_build_sensitivity`, etc.) must be
  defined BEFORE the `if submitted:` block
- Pipeline runs synchronously (no threading) — use `st.progress()` for UX
- `_secret(key)` helper: `st.secrets.get(key, "")` with `os.environ` fallback
- Clear `st.session_state["demo"]` after successful run

### Pipeline step output keys (critical — don't get these wrong)
| Step | Key in results_dict | Notes |
|------|--------------------|----|
| Market research | `market_data` | Dict with demographics, cap_rates, crime, walkscore, hud_fmr, rentcast, flood_zone |
| Rent prediction | `rent_prediction` | Keys: predicted_rates, predicted_rents_per_unit, predicted_annual_revenue, avg_predicted_growth |
| Pro forma | `results` | Keys: metrics, reversion, pro_forma, annual_btcfs, amortization, inputs |
| ML valuation | `ml_valuation` | Keys: assessment, predicted_total_value, actual_price_per_unit, feature_importances, premium_discount_pct |
| Lease analysis | `lease_analysis` | Keys: portfolio_summary, individual_leases, lease_count, input_comparison |
| Monte Carlo | `monte_carlo` | Keys: mean_irr, std_irr, percentiles (dict), probabilities, histogram_bins, histogram_counts — IRR already in % |
| Scenarios | `scenarios` | Keys: base, bull, bear (flat dicts), comparison_table (list of rows), bull_offsets, bear_offsets |
| Waterfall | `waterfall` | Keys: lp_irr, gp_irr, lp_total_return, gp_total_return, tiers |
| AI memo | `ai_memo` | Keys: executive_summary, market_analysis, financial_analysis, risk_factors, investment_thesis, recommendation |

---

## Milestone 2: Results Dashboard (`pages/1_Results.py`)

### What to build
11-tab results page reading from `st.session_state["results"]` with a what-if
sidebar that recalculates returns in real time.

### Critical key mappings (bugs to avoid)
- **ML guard:** `ml.get("assessment")` — NOT `ml.get("available")` (doesn't exist)
- **ML values:** `predicted_total_value`, `actual_price_per_unit` — NOT `predicted_value`, `actual_price`
- **ML features:** `list(ml.get("feature_importances", {}).keys())` — NOT `top_features`
- **MC guard:** `mc.get("mean_irr")` — NOT `mc.get("irr_mean")`
- **MC percentiles:** `mc["percentiles"]["P10"]`, `mc["percentiles"]["P50 (Median)"]`, `mc["percentiles"]["P90"]`
- **MC display:** values already in % — use `f"{v:.1f}%"` NOT `_pct(v)` (which multiplies by 100)
- **MC chart:** use `go.Bar(x=histogram_bins, y=histogram_counts)` — NOT `px.histogram(irr_samples)`
- **Rent table:** build from `predicted_rates` + `predicted_rents_per_unit` — NOT `yearly_predictions`
- **Scenarios table:** use `scenarios["comparison_table"]` directly — NOT iterating all dict items
- **Scenarios chart:** iterate only `["base", "bull", "bear"]` keys — skip comparison_table/offsets

### What-if sidebar
```python
sim_deal = copy.deepcopy(deal_obj)
sim_deal.revenue_growth_rate = wi_rent_growth / 100
wi_result = build_pro_forma(sim_deal)
wi_metrics = wi_result.get("metrics", metrics)
```

---

## Milestone 3: AI Deal Chat (`pages/2_Chat.py`)

### What to build
Claude claude-opus-4-6 chat grounded in deal results. System prompt includes all key numbers.

### Key notes
- Use `ANTHROPIC_API_KEY` (not OpenAI) for chat
- Anthropic API: `system=` is a separate param, not part of messages list
- MC keys in system prompt: `mc.get("mean_irr")`, `mc["percentiles"]["P10"]` etc.
- ML keys in system prompt: `ml.get("assessment")`, `ml.get("predicted_total_value")`
- Keep last 10 turns only: `st.session_state["chat_history"][-10:]`

---

## Milestone 4: Deal Comparison (`pages/3_Compare.py`)

### What to build
Side-by-side comparison of saved deals with Plotly charts and green highlighting.

### Key notes
- Read from `st.session_state["saved_deals"]` (dict keyed by job_id)
- Use `df.style.apply(_highlight_best, axis=0)` for column-wise green highlighting
- Guard: `st.stop()` if no saved deals

---

## Milestone 5: Report Export

### What to build
Excel/Word/PDF generators that receive the full pipeline output dict and write files
to `OUTPUT_DIR` (from `config.py`).

### Key notes
- All three generators called in `streamlit_app.py` after main pipeline
- Generators wrapped in try/except — failures store `None` path, shown gracefully
- Download buttons check `os.path.exists(path)` before rendering
- Files named `underwriting_{job_id}.xlsx` etc. using 12-char hex UUID

---

## Deployment (Streamlit Community Cloud)

1. Push to GitHub (main file: `streamlit_app.py`)
2. Connect repo at share.streamlit.io
3. Add all 7 API keys in the Streamlit Cloud Secrets editor (TOML format)
4. App auto-redeploys on every push to master
