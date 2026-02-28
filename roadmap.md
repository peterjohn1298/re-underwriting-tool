# Roadmap — RE Underwriting Intelligence Platform

Building on: Flask + numpy-financial + scikit-learn + FRED/Census/RentCast/HUD/WalkScore APIs + Claude + Gemini

---

## Completed Sections ✅

### ✅ 1. Core Financial Engine
10-year pro forma, IRR/DSCR/equity multiple/CoC/yield-on-cost, amortization schedule,
tax analysis with depreciation, sensitivity tables (4×), Monte Carlo simulation (1000 runs).

### ✅ 2. Market Data Pipeline
FRED macro signals, Census ACS demographics, BLS employment, Zillow ZORI rent index,
RentCast live comps + AVM, HUD Fair Market Rents, FBI UCR crime risk, Walk Score location scores.

### ✅ 3. ML / Analytical Layer
Polynomial regression rent predictor (FRED+ZORI blended), backtesting with time-series split,
GradientBoosting property valuation (19 features), multi-signal recommendation engine
(IRR + DSCR + ML + lease + Monte Carlo → BUY/HOLD/PASS score).

### ✅ 4. Lease Document Analysis
Multi-PDF upload, Gemini/Claude NLP extraction (tenant, rent, escalations, clauses, risk flags),
portfolio aggregation across multiple leases, weighted average escalation calculation.

### ✅ 5. Report Generation + CI/CD
Excel (pro forma + sensitivity + Monte Carlo + market data), Word investment memo,
PDF report, GitHub Actions CI/CD (lint + 51 tests + auto-deploy to Render), DRIVER plugin.

---

## Completed Sections (continued) ✅

### ✅ 6. AI Chat Interface (Deal Assistant)
Claude-powered conversational Q&A grounded in the complete deal analysis results.
Analyst can ask: "What's the main risk?", "What if vacancy hits 10%?", "How does
this compare to market FMR?" — all answered in context of the actual deal data.

### ✅ 7. AI-Generated Deal Memo
Claude (claude-sonnet-4-6) writes 6-section investment narrative grounded in deal data:
executive summary, market analysis, property overview, financial highlights, risk factors,
recommendation. Auto-generated during pipeline, embedded in Word doc, displayed on results
page with on-demand regeneration button. No figures fabricated.

### ✅ 8. Unit Mix Modeling
Studio/1BR/2BR/3BR breakdown with individual unit counts, in-place rents, and market rents.
Blended rent auto-overrides the single average. HUD FMR comparison per bedroom type.
Unit mix table on results page + Word doc. Live revenue calculator in the form.

---

### ✅ 9. Waterfall / LP-GP Promote
Two-tier promote waterfall: return of capital → LP preferred return (8% compounded) →
residual split (LP/GP promote). Outputs LP IRR, GP IRR, LP equity multiple, GP equity
multiple, promote dollars. Form toggle with configurable LP%, pref return%, promote%.

### ✅ 10. Scenario Comparison (Base / Bull / Bear)
Auto-generated Base/Bull/Bear scenarios with configurable rent growth, occupancy,
and exit cap offsets. Side-by-side comparison table on results page. Runs without
user configuration — default offsets: Bull +1.5% growth/+3% occ/−25bps exit,
Bear −1.5%/−5%/+50bps.

### ✅ 11. OpenFEMA Flood Zone + Database Persistence
FEMA NFHL REST API flood zone lookup by lat/lon (no API key). Identifies Zone AE, X, V,
etc. with SFHA flag and lender insurance requirement. SQLite persistence via services/database.py
— deals saved on completion, survive server restarts, reloadable via /api/reload_deal/<id>.

---

## Dependencies

```
S1 [Core Financial Engine] ──► S2 [Market Data Pipeline]
S1 ──────────────────────────► S3 [ML / Analytical Layer]
S1 ──────────────────────────► S4 [Lease Document Analysis]
S2 ──► S3 ──► S4 ──► S5 [Report Generation + CI/CD]   ✅ DONE
S5 ──► S6 [AI Chat Interface]
S5 ──► S7 [AI-Generated Deal Memo]
S1 ──► S8 [Unit Mix Modeling]
S8 ──► S9 [Waterfall / LP-GP]
S1 ──► S10 [Scenario Comparison]
S2 ──► S11 [Flood Zone + Persistence]
```

Build order for remaining work: S6 → S7 → S8 → S9 → S10 → S11

---

## DRIVER Stage Tracker

| Stage | Status | Artifact |
|-------|--------|----------|
| Define (开题调研) | ✅ Complete | `product-overview.md` |
| Represent (Roadmap) | ✅ Complete | `roadmap.md` (this file) |
| Implement | ✅ Complete | `app.py`, `models/`, `services/` — 11/11 sections |
| Validate | ✅ Active | `validation.md` (152 tests, CI green, manual QA all sections) |
| Evolve | ✅ Active | Deployed on Render, CI/CD auto-deploy on master push |
| Reflect | ✅ Active | `reflect.md` — updated with professor feedback iteration |

---

_One thing to expect: this plan will be wrong in some way. Implementation always
surfaces things planning can't. When that happens, update this file. The R↔I loop
is how real work gets done._
