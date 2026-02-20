# Section Extension Prompt Template

Use this template when adding a new feature or section to the platform.

---

I'm extending the RE Underwriting Intelligence Platform (Python + Streamlit).

**What I want to add:** [DESCRIBE YOUR FEATURE HERE]

**Where it fits:**
- [ ] New tab on the Results page (`pages/1_Results.py`)
- [ ] New pipeline step in `streamlit_app.py` (runs during analysis)
- [ ] New model in `models/` (pure Python, no UI)
- [ ] New service/API client in `services/api_clients/`
- [ ] New page in `pages/`

**Context:**
- Session results dict: `st.session_state["results"]` — add your output here with a new key
- API keys: add to `.streamlit/secrets.toml` and retrieve via `st.secrets.get("YOUR_KEY", "")`
- All pipeline steps run synchronously in `streamlit_app.py` with `st.progress()` updates
- Pure model functions go in `models/` with no Streamlit imports

**Relevant existing files:**
- `models/assumptions.py` — DealInputs dataclass (add fields here if needed)
- `models/financial_model.py` — Core pro forma engine
- `services/market_research.py` — Orchestrates all API calls
- `pages/1_Results.py` — Results dashboard (add new tab or section here)

Please implement this following the existing code style and patterns.
Do not modify models or services that aren't related to this feature.
