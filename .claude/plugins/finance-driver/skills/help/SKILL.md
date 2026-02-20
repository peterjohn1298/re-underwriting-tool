---
name: help
description: DRIVER overview, available skills, and philosophy explanation
---

# DRIVER Help

## What is DRIVER?

DRIVER is a methodology for building finance and quantitative analysis tools with AI assistance. It guides you from concept to completion through six stages.

---

## The Philosophy: Cognition Mate (认知伙伴)

**Core principle:** 互帮互助，因缘合和，互相成就

| Chinese | Pinyin | Meaning |
|---------|--------|---------|
| 互帮互助 | hu bang hu zhu | Mutual help, helping each other |
| 因缘合和 | yin yuan he he | Causes and conditions coming together (interdependent arising) |
| 互相成就 | hu xiang cheng jiu | Accomplishing together, mutual achievement |

**What this means in practice:**
- AI is not a tool you command — it's a thinking partner
- You bring vision and domain expertise; AI brings patterns and research
- Neither creates alone — meaning emerges from interaction
- The relationship is collaborative, not transactional

---

## The Six Stages

```
DEFINE (开题调研)
    "What are we building? What already exists?"
    开题 = Open the topic  调研 = Research/investigate
                       ↓
REPRESENT
    "How do we break this into buildable pieces?"
    Roadmap, data model, sections
                       ↓
IMPLEMENT
    "Build it, run it, show it"
    Show don't tell — code speaks louder than plans
                       ↓
VALIDATE
    "Cross-check your instruments"
    Known answers, reasonableness, edges, AI blind spots
                       ↓
EVOLVE
    "Package the final deliverable"
    Self-contained export, ready for production
                       ↓
REFLECT (Optional)
    "What did we learn?"
    Capture tech stack lessons, especially failures
```

---

## Project Structure

DRIVER uses a flat, easy-to-navigate file structure. All project files live in a single folder (default: `my-project/`, customizable via `.driver.json`):

```
.driver.json                      # Project config (folder name)
my-project/
├── README.md                     # Project overview and structure
├── research.md                   # Research findings (分头研究)
├── product-overview.md           # Product definition (your PRD)
├── roadmap.md                    # 3-5 buildable sections
├── data-model.md                 # Core entities and relationships
├── validation.md                 # Cross-check results (all sections)
├── reflect.md                    # Learnings and retrospective
├── spec-[section-name].md        # Section specifications
├── design/                       # Web apps only
│   ├── tokens.json               # Colors and typography
│   └── shell.md                  # Navigation shell spec
└── build/                        # Implementation artifacts
    └── [section-id]/
        ├── data.json             # Sample data
        └── types.ts              # TypeScript interfaces
```

Human-readable documents live at the project root. Implementation artifacts live in `build/`.

---

## Key Concepts

### 分头研究 (fen tou yan jiu)
**"Parallel research"** — Before building anything, research what exists. You focus on your unique needs; AI researches existing libraries, papers, implementations.

很可能已经有类似的了 = "There's probably something similar already"

### Show Don't Tell
Don't explain what you'll build. **Build it. Run it. Let them see it.**

The fastest feedback loop: See result → Give feedback → Iterate → See updated result

### KISS — Keep It Simple, Structured
- Simple and logical beats elegant and fancy
- Quants need clear data tables, not animations
- A 500-line Python script beats a 50-file TypeScript project

### Persistent Artifacts
Every DRIVER stage produces a markdown file — research.md, product-overview.md, roadmap.md, spec files. These are your **shared mutable state**:
- They survive context window limits (chat history gets compressed; files don't)
- They enable asynchronous review (read at your own pace, catch mistakes)
- They serve as review surfaces where you annotate corrections

**Rule:** If research, a plan, or a decision lives only in chat, it will get lost. Write it to a file.

### The Annotation Cycle
After AI writes a plan or spec, don't just say "looks good." Review it in your editor:
1. AI writes the plan/spec to a markdown file
2. You open it and add inline notes (corrections, domain knowledge, rejected approaches)
3. Send it back: "Update based on my annotations — don't implement yet"
4. AI revises the document
5. Repeat 1-6 times until the plan is right

**This is where the real creative work happens.** Implementation should be mechanical.

### Mermaid Diagrams
DRIVER uses Mermaid diagrams as standard visual documentation:
- **System context** — in product-overview.md (how the tool fits in the workflow)
- **Dependencies** — in roadmap.md (build order between sections)
- **User flows** — in spec files (step-by-step interaction)
- **ER diagrams** — in data-model.md (entity relationships)
- **Decision landscapes** — in research.md (comparing options)

---

## Available Skills

| Skill | Stage | Purpose |
|-------|-------|---------|
| `/finance-driver:init` | Setup | Initialize project structure |
| `/finance-driver:status` | Any | Check progress, get suggestions |
| `/finance-driver:help` | Any | This help page |
| `/finance-driver:research` | Any | Lightweight 分头研究 — research libraries, approaches, references |
| `/finance-driver:define` | DEFINE | Research and define vision |
| `/finance-driver:represent-roadmap` | REPRESENT | Break into sections |
| `/finance-driver:represent-datamodel` | REPRESENT | Define core entities |
| `/finance-driver:represent-tokens` | REPRESENT | Colors/typography (web apps) |
| `/finance-driver:represent-shell` | REPRESENT | Navigation shell (web apps) |
| `/finance-driver:represent-section` | REPRESENT | Spec a section |
| `/finance-driver:implement-data` | IMPLEMENT | Sample data (web apps) |
| `/finance-driver:implement-screen` | IMPLEMENT | Build and run code |
| `/finance-driver:validate` | VALIDATE | Cross-check: known answers, reasonableness, edges, AI risks |
| `/finance-driver:evolve` | EVOLVE | Generate export package |
| `/finance-driver:reflect` | REFLECT | Capture learnings |

---

## Recommended Stack for Finance/Quant

```
UI:           Streamlit (or Dash/Panel)
Backend:      FastAPI + Pydantic
Calculations: NumPy, Pandas, SciPy
Finance:      numpy-financial, QuantLib
Data Sources: See README for tiered recommendations
              LLM-Native: financialdatasets.ai, Alpha Vantage, EODHD
              MCP Available: Polygon.io, S&P Global/Kensho
              Free (verify): yfinance, FRED
Storage:      SQLite → PostgreSQL, Parquet files
Testing:      pytest + Hypothesis
```

> **Data Quality Matters:** For LLM-driven development, use MCP-native data providers (financialdatasets.ai recommended). Free sources like yfinance may have gaps, delays, or inaccuracies.

**Why Python over TypeScript for quant work:**
- Vectorized calculations (NumPy) vs manual loops
- Pydantic catches validation errors at boundaries
- `streamlit run app.py` vs npm/webpack complexity
- Division by zero: `np.divide(..., where=b!=0)` vs manual guards everywhere

---

## Example Projects

| Project | Style | Key Libraries | Data Source |
|---------|-------|---------------|-------------|
| DCF Valuation Tool | Damodaran | numpy-financial | financialdatasets.ai |
| Portfolio Optimizer | Markowitz | PyPortfolioOpt, scipy.optimize | Professional feed |
| Factor Research | Open Source AP | pandas, statsmodels, alphalens | WRDS, CRSP |
| Risk Dashboard | VaR/CVaR | scipy.stats, matplotlib | Professional feed |
| Data Pipeline | ETL | pandas, SQLAlchemy | Multiple sources |

---

## Getting Started

1. **New project:** `/finance-driver:init` or just describe what you want to build
2. **Existing project:** `/finance-driver:status` to see where you are
3. **Stuck?** Tell me the finance problem you're solving — we'll figure it out together

---

## Iron Laws (Never Break These)

| Stage | Iron Law |
|-------|----------|
| DEFINE | NO BUILDING WITHOUT 分头研究 FIRST |
| REPRESENT | PLAN THE UNIQUE PART — DON'T REINVENT |
| IMPLEMENT | SHOW DON'T TELL — BUILD AND RUN IT |
| VALIDATE | CROSS-CHECK YOUR INSTRUMENTS — four checks, every time |
| EVOLVE | SELF-CONTAINED DELIVERABLE |
| REFLECT | CAPTURE WHAT DIDN'T WORK |
