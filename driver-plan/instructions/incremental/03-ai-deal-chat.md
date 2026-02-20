# Milestone 3: AI Deal Chat

**File:** `pages/2_Chat.py`
**Depends on:** Milestone 1 (session state), ANTHROPIC_API_KEY

## Goal
Claude-powered deal analyst grounded in analysis results.

## Acceptance Criteria
- [ ] Guard redirects if no results in session state
- [ ] Warning shown if ANTHROPIC_API_KEY not configured
- [ ] System prompt includes all key metrics, market data, ML, MC, recommendation
- [ ] Claude responds only from deal data (no hallucinated numbers)
- [ ] Last 10 conversation turns maintained in session state
- [ ] "🗑️ Clear Chat" button clears history

## Key Pitfalls
- Anthropic API takes `system=` as separate param, NOT inside messages list
- Use `mc.get("mean_irr")` and `mc["percentiles"]["P10"]` in system prompt (not old irr_mean/irr_p10 keys)
- Use `ml.get("assessment")` and `ml.get("predicted_total_value")` (not available/predicted_value)
