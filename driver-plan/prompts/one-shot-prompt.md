# One-Shot Prompt — RE Underwriting Platform

Paste this into any coding agent to rebuild or significantly extend the platform.

---

You are helping me build/extend the RE Underwriting Intelligence Platform — an
institutional-grade multifamily real estate underwriting tool deployed on Streamlit
Community Cloud.

Before starting, please answer these questions:

1. **Mode:** Are we (a) rebuilding from scratch, (b) adding a new feature, or (c) fixing a bug?
2. **Stack constraint:** Python + Streamlit only. No React, no TypeScript, no Node.
3. **API keys available:** FRED, RentCast, HUD, Walk Score, FBI, FEMA, OpenAI (GPT-4o), Anthropic (Claude claude-opus-4-6), Google (Gemini).
4. **Deployment target:** Streamlit Community Cloud (secrets via `st.secrets`).

**What's already built:**
- 5 pages: input form, 11-tab results dashboard, AI chat, deal comparison, report export
- 13-step analysis pipeline: market research → pro forma → ML valuation → Monte Carlo → scenarios → waterfall → AI memo → Excel/Word/PDF
- All models in `models/` are pure Python with zero UI dependency

**Key architectural rules:**
- Session state key: `st.session_state["results"]` holds the full pipeline output dict
- API key resolution: `st.secrets.get(key, "")` with `os.environ` fallback
- Monte Carlo IRR values are already in % — never multiply by 100 again
- Scenarios output: flat keys per scenario + `comparison_table` list (not nested `metrics` dict)
- ML output guard: `ml.get("assessment")` not `ml.get("available")`

Refer to `driver-plan/product-overview.md` and `driver-plan/instructions/one-shot-instructions.md`
for full context before writing any code.
