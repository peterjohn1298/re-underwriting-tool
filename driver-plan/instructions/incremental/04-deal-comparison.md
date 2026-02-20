# Milestone 4: Deal Comparison

**File:** `pages/3_Compare.py`
**Depends on:** Milestone 1 + 2 (saved deals in session state)

## Goal
Side-by-side comparison of multiple saved deals.

## Acceptance Criteria
- [ ] Guard shows "no deals saved" message if session_state["saved_deals"] empty
- [ ] Comparison table with color highlighting for best IRR/EM/CoC
- [ ] NOI growth line chart with one line per deal
- [ ] Remove selected deals and rerun

## Key Notes
- Deals saved as dict keyed by job_id in session_state["saved_deals"]
- Use `df.style.apply(_highlight_best, axis=0)` for column-wise highlighting
- Parse string values (strip %, x, $, commas) to find numeric max for highlighting
