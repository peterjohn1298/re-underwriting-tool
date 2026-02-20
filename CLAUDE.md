# RE Underwriting Intelligence Platform — Project Instructions

## Project Overview

Flask-based commercial real estate underwriting tool. All 11 DRIVER sections complete and deployed on Render.

See `product-overview.md` for full context, `roadmap.md` for section list, `validation.md` for QA evidence, `reflect.md` for lessons learned.

---

## PAL MCP Integration

This project uses PAL MCP tools at specific DRIVER stages. Use them **proactively** — don't wait to be asked.

### Tool → Stage Map

| DRIVER Stage | PAL Tool | When to Trigger |
|---|---|---|
| Implement | `mcp__pal__debug` | Any bug not solved in 2 attempts |
| Implement | `mcp__pal__thinkdeep` | Architectural decision with >2 valid approaches |
| Implement | `mcp__pal__apilookup` | Before adding or changing any external API integration |
| Implement | `mcp__pal__chat` | Complex financial logic that needs a second opinion |
| Validate | `mcp__pal__codereview` | After every new or modified section — before marking complete |
| Validate | `mcp__pal__precommit` | Before every commit to master (CI auto-deploys to Render) |
| Evolve | `mcp__pal__consensus` | Major architectural decisions (React migration, real ML data) |
| Any | `mcp__pal__chat` | Collaborative thinking on ambiguous requirements |

### Finance-Specific Rules

1. **Run `mcp__pal__codereview`** after any change to financial calculations (IRR, DSCR, Monte Carlo, waterfall promote). Focus: formula correctness, zero/negative edge cases, off-by-one in amortization periods.

2. **Run `mcp__pal__precommit`** before every push to master. CI auto-deploys to Render — a broken commit goes live immediately.

3. **Use `mcp__pal__apilookup`** before modifying FRED, RentCast, HUD, or Census API integrations. These APIs have undocumented behaviors — see `reflect.md` for known issues.

4. **Use `mcp__pal__debug`** when market data returns unexpected values. Root cause is almost always in the API response structure (key name, nested dict), not the calling code.

5. **Use `mcp__pal__thinkdeep`** for the open architectural questions in `reflect.md`: React migration, real transaction data for ML training, DB-first persistence.

6. **Use `mcp__pal__chat` with `model: "gemini-2.5-pro"`** for complex finance domain questions — waterfall mechanics, FMR matching logic, Monte Carlo parameter calibration.

---

## Key Files

| File | Purpose |
|---|---|
| `app.py` | Flask routes, background job orchestration, chat API |
| `models/financial_model.py` | Core calculations: IRR, DSCR, amortization, equity multiple |
| `models/metrics.py` | Monte Carlo, sensitivity tables, scenario comparison |
| `services/market_research.py` | Parallel API pipeline: FRED, RentCast, HUD, Census, Walk Score |
| `services/ai_service.py` | Claude + Gemini: lease extraction, deal memo, chat grounding |
| `services/waterfall.py` | LP/GP promote waterfall calculation |
| `services/database.py` | SQLite deal persistence |
| `services/flood_zone.py` | FEMA NFHL flood zone lookup |
| `tests/` | 51 pytest tests — always run before committing |
| `validation.md` | Per-section QA evidence |
| `reflect.md` | What was built, what the plan got wrong |

---

## Workflow Rules

- Run `pytest tests/` before any commit — 51 tests must pass
- Run `flake8 --select=E9,F63,F7,F82 .` before commit (CI fails on these error codes)
- Never commit API keys — use environment variables only (ANTHROPIC_API_KEY, RENTCAST_API_KEY, etc.)
- ML valuation uses synthetic training data — always label outputs as "synthetic_calibrated"
- All AI-generated content (memo, chat) must be grounded in deal data — no fabricated numbers

---

## PAL codereview Default Settings for This Project

```
review_type: "full"
focus_on: "financial formula correctness, API error handling, AI grounding safety"
severity_filter: "high"
model: "gemini-2.5-pro"
thinking_mode: "high"
```

## PAL precommit Default Settings for This Project

```
path: "/mnt/c/ai project folder/re-underwriting-tool-clone"
focus_on: "broken tests, API key leaks, flake8 E9/F63/F7/F82 errors"
precommit_type: "external"
model: "gemini-2.5-pro"
```
