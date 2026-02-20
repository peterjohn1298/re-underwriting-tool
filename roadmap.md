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

## In-Progress Sections 🔨

### 6. AI Chat Interface (Deal Assistant)
Claude-powered conversational Q&A grounded in the complete deal analysis results.
Analyst can ask: "What's the main risk?", "What if vacancy hits 10%?", "How does
this compare to market FMR?" — all answered in context of the actual deal data.

### 7. AI-Generated Deal Memo
Claude writes the full investment memorandum using real market data as context —
replacing the current fill-in template. Sections: executive summary, market analysis,
property description, financial summary, risks, recommendation.

---

## Planned Sections 📋

### 8. Unit Mix Modeling
Detailed breakdown by bedroom type (Studio/1BR/2BR/3BR) with individual unit counts,
rents, and vacancy rates. Feeds the financial model with bedroom-weighted blended rent
instead of a single average. Required for accurate HUD FMR comparison per bedroom type.

### 9. Waterfall / LP-GP Promote
LP/GP capital structure with preferred return, catch-up, and promote tiers.
Models: who gets what, when, under Base/Bull/Bear scenarios. Essential for
institutional and fund-level underwriting.

### 10. Scenario Comparison (Base / Bull / Bear)
Run three scenarios simultaneously with different rent growth, vacancy, and exit cap
assumptions. Side-by-side comparison dashboard. Analyst picks the scenario; tool shows
the distribution of outcomes.

### 11. OpenFEMA Flood Zone + Database Persistence
Flood zone designation by address (NFIP via OpenFEMA — free, no key).
SQLite database so deals persist across server restarts with versioning history.

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
| Implement | 🔨 In Progress | `app.py`, `models/`, `services/` |
| Validate | ✅ Partial | `validation.md` (51 tests passing, CI live) |
| Evolve | 🔨 In Progress | Deployed on Render, CI/CD active |
| Reflect | ⏳ Pending | `reflect.md` (after all sections complete) |

---

_One thing to expect: this plan will be wrong in some way. Implementation always
surfaces things planning can't. When that happens, update this file. The R↔I loop
is how real work gets done._
