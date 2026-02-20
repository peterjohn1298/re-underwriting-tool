# Milestone 1: Deal Input & Analysis Pipeline

**File:** `streamlit_app.py`
**Depends on:** All models/ and services/ (pre-built)

## Goal
Multi-tab input form + full 13-step analysis pipeline that stores results in
session state and navigates to the Results Dashboard.

## Acceptance Criteria
- [ ] Form renders with 4 tabs: Deal Sheet, Unit Mix, AI/ML Features, Override Defaults
- [ ] "📋 Load Demo Deal" fills all fields including unit mix (keyed widgets)
- [ ] ML + Rent Prediction checkboxes are auto-enabled by demo button
- [ ] Validation blocks submit if address, price, or rent missing
- [ ] Progress bar advances with descriptive labels for each of 13 steps
- [ ] All pipeline outputs stored in `st.session_state["results"]` with correct keys
- [ ] Navigates to `pages/1_Results.py` on success

## Key Pitfalls
- Keyed widgets (`key=f"um_count_{ut}"`) ignore `value=`; set `st.session_state[key]` directly in demo handler
- Helper functions must be defined before `if submitted:` block
- `_secret(key)` helper needed for Streamlit Cloud vs local dev API key resolution
