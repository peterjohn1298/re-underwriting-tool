# RE Underwriting Intelligence Platform — Reflections

## Project Summary

**Product:** RE Underwriting Intelligence Platform
**Sections:** 5 (Input Pipeline, Results Dashboard, AI Chat, Deal Comparison, Report Export)
**Tech Stack Used:** Python + Streamlit, scikit-learn, Plotly, 7 external APIs, OpenAI + Anthropic + Google
**Deployment:** Streamlit Community Cloud (public URL, zero infrastructure)

---

## What Worked Well

- **95% backend reuse** — All of `models/` and `services/` had zero Flask dependency.
  The Streamlit conversion only touched the UI layer; all financial logic ran untouched.
- **Synchronous pipeline in Streamlit** — No threading, no background jobs dict, no polling.
  `st.progress()` + sequential function calls was simpler and more debuggable than Flask's
  background job pattern.
- **Session state as data bus** — `st.session_state["results"]` cleanly replaced Flask's
  in-memory `jobs` dict for passing pipeline output between pages.
- **Streamlit Community Cloud deployment** — Zero server config, auto-deploys on git push,
  free public URL. Exactly right for sharing with a professor.
- **`_secret()` helper pattern** — `st.secrets.get(key, "")` with `os.environ` fallback
  worked cleanly for both Streamlit Cloud and local `.env` development.
- **Demo button** — "📋 Load Demo Deal" significantly improved usability for demos and testing.

---

## Challenges & Learnings

### 1. Streamlit Keyed Widgets Ignore `value=`
**Problem:** Unit mix number inputs had explicit `key=f"um_count_{ut}"`. The demo button
set `st.session_state["demo"]` and called `st.rerun()`, expecting `value=d.get(...)` to
populate them. It didn't — keyed widgets are controlled entirely by session state, not `value=`.

**Lesson:** For any widget with an explicit `key=`, always set `st.session_state[key]`
directly to change its value programmatically. The `value=` parameter is only the initial
default on first render.

```python
# WRONG — value= ignored for keyed widgets after first render
st.number_input("# Units", value=d.get("um_count_1BR", 0), key="um_count_1BR")

# RIGHT — set session state directly before rerun
st.session_state["um_count_1BR"] = DEMO_VALUES["um_count_1BR"]
st.rerun()
```

### 2. Model Output Key Mismatches (Multiple Bugs)
**Problem:** The Results page was written with assumed key names that didn't match what
the models actually returned. Every ML, Monte Carlo, and Rent Forecast section was broken.

**Bugs found:**
| Wrong key assumed | Actual key | Model |
|-------------------|-----------|-------|
| `irr_mean` | `mean_irr` | monte_carlo.py |
| `irr_std` | `std_irr` | monte_carlo.py |
| `irr_p10/p50/p90` | `percentiles["P10"]` etc. | monte_carlo.py |
| `irr_samples` | `histogram_bins` + `histogram_counts` | monte_carlo.py |
| `ml.get("available")` | `ml.get("assessment")` | ml_valuation.py |
| `predicted_value` | `predicted_total_value` | ml_valuation.py |
| `actual_price` | `actual_price_per_unit` | ml_valuation.py |
| `top_features` (list) | `feature_importances` (dict keys) | ml_valuation.py |
| `yearly_predictions` | `predicted_rates` + `predicted_rents_per_unit` | rent_predictor.py |

**Lesson:** Always read the model source to confirm output keys before writing display code.
Never assume key names — they are often singular vs plural, nested vs flat, or named
differently than intuition suggests.

### 3. Monte Carlo Values Already in Percentage
**Problem:** The Results page used `_pct(mc.get("mean_irr"))` which multiplied by 100.
But `mean_irr` is already stored as a percentage (e.g., `14.2`, not `0.142`), so the
display showed `1420%` instead of `14.2%`.

**Lesson:** Always check whether a model returns ratios (0–1) or percentages (0–100).
Monte Carlo stores IRR as % directly. Pro forma metrics store IRR as ratio (0–1).

### 4. Scenarios Dict Iteration Bug
**Problem:** The Results page did `for scenario_name, sc_data in scenarios.items()` which
iterated ALL keys including `comparison_table` (a list), `bull_offsets`, `bear_offsets`.
It also looked for `sc_data.get("metrics", {})` — a nested sub-dict that doesn't exist
(data is flat).

**Lesson:** When a model returns a mixed dict (some values are scenario dicts, some are
metadata), always iterate specific keys explicitly (`["base", "bull", "bear"]`) rather than
all items. Check the model source to confirm exact structure.

### 5. TOML Secrets Formatting on Streamlit Cloud
**Problem:** Pasting the HUD JWT token (390 chars) into Streamlit Cloud's secrets editor
caused "Invalid format: please enter valid TOML" — hidden newlines were inserted during
copy-paste, breaking the string.

**Lesson:** For long token values, write the secrets.toml file directly (using the Write
tool) and have the user copy from a local text editor (Notepad), not from a rendered
markdown block. Always verify TOML validity with a TOML parser before saving.

### 6. ANTHROPIC_API_KEY Was Configured But Never Used
**Problem:** The Anthropic key was in secrets.toml and config.py, but the AI chat was
using OpenAI GPT-4o. The key sat unused.

**Lesson:** Audit all configured API keys against actual usage at build time. If a key is
configured, it should be used — or removed to avoid confusion.

---

## Tech Stack Retrospective

### What Worked
- **Streamlit** — Right choice for a shareable single-URL analytics tool. Fast iteration,
  no frontend build step, native Python types throughout.
- **Plotly** — Clean integration with `st.plotly_chart()`, dark theme support via
  `paper_bgcolor`/`plot_bgcolor`, flexible for histograms, bar charts, line charts.
- **scikit-learn** — GradientBoostingRegressor for ML valuation worked well; Ridge for
  rent prediction was adequate given limited training data.
- **Anthropic Claude claude-opus-4-6** — Correct choice for the grounded deal chat; system
  prompt approach (separate from messages) is clean and effective.

### What Caused Friction
- **Streamlit session state with keyed widgets** — Counter-intuitive behavior; `value=`
  is ignored for widgets with explicit keys. Required direct session_state assignment.
- **Streamlit multi-page navigation** — `st.switch_page()` works but the sidebar auto-shows
  all pages by filename; required CSS or `initial_sidebar_state="collapsed"` management.
- **PDF generation on Streamlit Cloud** — reportlab works but WeasyPrint/wkhtmltopdf would
  need system-level dependencies not available on Cloud free tier.

### Next Time, Use Instead
| Used | Consider Instead | Why |
|------|-----------------|-----|
| Manual key mapping | Read model source first | Prevents all key mismatch bugs |
| Copy-paste for TOML | Write file directly | Avoids hidden newline corruption |
| OpenAI for chat | Anthropic (already configured) | Use the key you have |

---

## Time Analysis

**Time wasted on:** Debugging key mismatches between model outputs and display code —
every ML, Monte Carlo, and Rent Forecast bug came from assuming key names instead of
reading the source. This was the single biggest source of rework.

**Would have saved time:** Reading all model output structures (monte_carlo.py,
ml_valuation.py, rent_predictor.py, scenarios.py) before writing the Results page display
code. 30 minutes of reading would have saved hours of debugging.

---

## Reusable Patterns

### Session State as Data Bus (Streamlit Multi-Page)
Store all pipeline output in a single dict key; pages read from it with guard clauses.
```python
st.session_state["results"] = results_dict  # writer page
R = st.session_state.get("results")         # reader page
if not R: st.switch_page("streamlit_app.py"); st.stop()
```

### API Key Resolution (Cloud + Local)
```python
def _secret(key, fallback=""):
    try:
        return st.secrets.get(key, fallback) or fallback
    except Exception:
        return os.environ.get(key, fallback)
```

### Keyed Widget Demo Population
```python
# In demo button handler — set state before rerun
for key, val in DEMO_VALUES.items():
    if key.startswith("um_"):  # keyed widgets
        st.session_state[key] = val
st.rerun()
```

### Synchronous Pipeline with Progress
```python
progress = st.progress(0, text="Starting...")
progress.progress(10, text="Fetching market data...")
result = run_step()
progress.progress(30, text="Building pro forma...")
```

---

## Libraries & Tools to Remember

### Always Use
- `copy.deepcopy(deal_obj)` for what-if scenarios — never mutate the original
- `plotly.subplots.make_subplots` for side-by-side charts
- `pandas.DataFrame.style.apply()` for conditional cell highlighting

### Avoid
- `px.histogram(x=raw_samples)` when you already have binned data — use `go.Bar(x=bins, y=counts)`
- Iterating `dict.items()` when only a subset of keys are scenario dicts

---

## Notes for Future Projects

- Read ALL model output structures before writing display code
- Audit all configured API keys — if configured, use it or remove it
- Test the demo button flow end-to-end before every deployment
- Streamlit Cloud free tier sleeps after inactivity — warn users to expect a ~30s cold start
- For TOML secrets with long values (JWTs, long API keys), write the file programmatically

---

*Captured using DRIVER*
