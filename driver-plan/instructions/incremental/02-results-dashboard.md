# Milestone 2: Results Dashboard

**File:** `pages/1_Results.py`
**Depends on:** Milestone 1 (session state populated)

## Goal
11-tab results page with what-if sidebar calculator and navigation buttons.

## Acceptance Criteria
- [ ] Guard redirects to input if no results in session state
- [ ] All 11 tabs render with correct data from session_state["results"]
- [ ] ML tab: shows valuation badge + predicted_total_value + feature_importances
- [ ] MC tab: histogram bar chart using histogram_bins/histogram_counts
- [ ] Rent tab: yearly table built from predicted_rates + predicted_rents_per_unit
- [ ] Scenarios tab: uses comparison_table from scenarios output
- [ ] What-if sidebar recalculates IRR/EM/CoC/DSCR in real time
- [ ] "💬 AI Chat" button navigates to pages/2_Chat.py
- [ ] "📌 Save for Comparison" saves deal to session_state["saved_deals"]

## Key Pitfalls
See `one-shot-instructions.md` → Milestone 2 → "Critical key mappings" section.
Every single one of those was a real bug that caused sections to not display.
