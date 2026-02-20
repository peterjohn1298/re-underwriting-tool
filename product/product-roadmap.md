# Roadmap

Building on: scikit-learn, numpy-financial, Streamlit, Plotly, 7 external APIs (FRED, RentCast,
HUD, Walk Score, FBI, FEMA, Census/BLS), OpenAI GPT-4o, Anthropic Claude claude-opus-4-6

## Sections

### 1. Deal Input & Analysis Pipeline
Multi-tab Streamlit form (Deal Sheet, Unit Mix, AI/ML Features, Override Defaults) that
collects ~50 deal parameters and runs the full underwriting pipeline — 7 API calls, 4 ML
models, Monte Carlo, scenarios, waterfall, AI memo, and Excel/Word/PDF export — in a single
synchronous flow with live progress updates.

### 2. Results Dashboard
11-tab results page (Overview, Returns, Pro Forma, Market Data, ML & Monte Carlo, Rent
Forecast, Scenarios, Lease Analysis, AI Memo, Waterfall, Downloads) showing every output of
the pipeline with interactive Plotly charts, metric cards, and a what-if sidebar calculator
that recalculates returns in real time when assumptions change.

### 3. AI Deal Chat
Claude claude-opus-4-6 powered chat assistant grounded in the deal's analysis results —
system prompt includes all financial metrics, market data, ML valuation, Monte Carlo
percentiles, and recommendation signals so Claude can answer deal-specific questions without
hallucinating numbers.

### 4. Deal Comparison
Side-by-side comparison page for multiple saved deals — color-highlighted dataframe (best
IRR/EM/CoC in green), NOI growth line chart across deals, and deal management (add/remove).

### 5. Report Export
One-click download of institutional-quality Excel workbook (pro forma + sensitivity +
market data + ML results), Word investment memo (narrative + comps + AI analysis), and
PDF summary report — all generated from the same pipeline output dict.
