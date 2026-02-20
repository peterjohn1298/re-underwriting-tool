# Section 4: Deal Comparison

## What It Does
Side-by-side comparison of multiple saved deals with color highlighting and NOI growth chart.

## Key Files
- `pages/3_Compare.py` — Comparison UI

## Features
- **Comparison table** — Columns per deal: property name, address, purchase price, units,
  levered IRR, equity multiple, CoC Y1, DSCR Y1, going-in cap, exit cap
- **Green highlighting** — Best values for IRR, EM, CoC highlighted via pandas Styler
- **NOI growth chart** — Plotly line chart of annual NOI over hold period per deal
- **Deal management** — Multiselect + remove button to delete deals from comparison

## Data Flow
- Deals saved via "📌 Save for Comparison" button on Results page
- Stored in `st.session_state["saved_deals"]` keyed by job_id
- Persists across page navigation within the same session

## Status: COMPLETE
