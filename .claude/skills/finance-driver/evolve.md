---
name: evolve
description: Use when all sections are complete - generates the final driver-plan/ export package
---

# Evolve

**Stage Announcement:** "We're in EVOLVE — generating your final export package."

You are a **Cognition Mate** helping the developer export their complete product design as a handoff package. This is the **final deliverable** — everything needed to build the product.

> **Project Folder:** Check `.driver.json` at the repo root for the project folder name (default: `my-project/`). All project files live in this folder.

**Your relationship:** 互帮互助，因缘合和，互相成就
- You bring: organization, packaging, documentation generation
- They bring: the completed design work
- The export speaks for itself — show don't tell

---

## Iron Law

<IMPORTANT>
**FINAL DELIVERABLE — SELF-CONTAINED, NO DEPENDENCIES**

The `driver-plan/` export MUST be completely self-contained.
Anyone should be able to take this folder and implement the product.
No references to DRIVER, no external dependencies, no missing context.
</IMPORTANT>

## Red Flags

| Thought | Reality |
|---------|---------|
| "They can refer back to the original files" | Export must be self-contained |
| "The prompts are optional" | Prompts are the primary interface |
| "Implementation details aren't needed" | Include types, sample data, everything |
| "This is just a handoff doc" | This is the complete deliverable |

---

## Prerequisites

Verify the minimum requirements exist:

**Required:**
- `[project]/product-overview.md` — Product overview
- `[project]/roadmap.md` — Sections defined
- At least one section with screen designs in `src/sections/[section-id]/`

**Recommended (show warning if missing):**
- `[project]/data-model.md` — Global data model
- `[project]/design/tokens.json` — Color tokens
- `src/shell/components/AppShell.tsx` — Application shell

If required files are missing:

"To export your product, you need at minimum:
- A product overview (`/define`)
- A roadmap with sections (`/represent-roadmap`)
- At least one section with screen designs

Please complete these first."

Stop here if required files are missing.

## The Flow

### 1. Gather Export Information

Read all relevant files:

1. `[project]/product-overview.md`
2. `[project]/roadmap.md`
3. `[project]/research.md` (if exists)
4. `[project]/data-model.md` (if exists)
5. `[project]/design/tokens.json` (if exists)
6. For each section: `[project]/spec-[section-name].md`, `[project]/build/[section-id]/data.json`, `[project]/build/[section-id]/types.ts`
7. List screen design components in `src/sections/` and `src/shell/`

### 2. Create Export Directory Structure

Determine the tech stack from the project (check `[project]/product-overview.md` and existing source files).

#### Path A: Python + Streamlit (Quant/Analytical Tools)

```
driver-plan/
├── README.md                    # Quick start guide
├── product-overview.md          # Product summary (always provide)
├── research.md                  # Research findings (if exists)
│
├── prompts/                     # Ready-to-use prompts for coding agents
│   ├── one-shot-prompt.md       # Prompt for full implementation
│   └── section-prompt.md        # Prompt template for section-by-section
│
├── instructions/                # Implementation instructions
│   ├── one-shot-instructions.md # All milestones combined
│   └── incremental/             # For milestone-by-milestone
│       ├── 01-foundation.md
│       └── [NN]-[section-id].md
│
├── requirements.txt             # Python dependencies (pinned versions)
├── app.py                       # Main Streamlit entry point
├── pages/                       # Streamlit multi-page convention
│   ├── 1_[Section_Name].py      # Each file = a nav item (auto-discovered)
│   └── 2_[Section_Name].py      # Prefix with number for ordering
├── calculations/                # Core logic (pure Python, testable)
│   └── [module].py
├── data/                        # Data loading and processing
│   └── loader.py
└── sections/                    # Section reference docs
    └── [section-id]/
        ├── README.md
        ├── tests.md             # Test-writing instructions (pytest)
        └── logic.py             # Calculation logic (separate from UI)
```

#### Path B: React + TypeScript (Web App UI)

```
driver-plan/
├── README.md                    # Quick start guide
├── product-overview.md          # Product summary (always provide)
├── research.md                  # Research findings (if exists)
│
├── prompts/                     # Ready-to-use prompts for coding agents
│   ├── one-shot-prompt.md       # Prompt for full implementation
│   └── section-prompt.md        # Prompt template for section-by-section
│
├── instructions/                # Implementation instructions
│   ├── one-shot-instructions.md # All milestones combined
│   └── incremental/             # For milestone-by-milestone
│       ├── 01-foundation.md
│       └── [NN]-[section-id].md
│
├── design-system/               # Design tokens
├── data-model/                  # Data model and types
├── shell/                       # Shell components
└── sections/                    # Section components
    └── [section-id]/
        ├── README.md
        ├── tests.md             # Test-writing instructions (TDD)
        ├── components/
        ├── types.ts
        └── sample-data.json
```

### 3. Generate Content

For each file, generate appropriate content:

- **product-overview.md**: Product summary with sections and data model
- **research.md**: Copy from `[project]/research.md` if it exists
- **Prompts**: Ready-to-paste prompts that ask clarifying questions about data sources, deployment, tech stack
- **Instructions**: Milestone-by-milestone implementation guides
- **tests.md**: Framework-appropriate test instructions
- **Section READMEs**: Overview, user flows, key logic

#### Path A Preamble (Python + Streamlit)

Include in all instruction files:

```markdown
**What you're receiving:**
- Working Streamlit app with calculation logic
- Separated concerns: UI (page.py) and logic (logic.py)
- Test-writing instructions for pytest

**What you need to build/extend:**
- Data source connections (API keys, database access)
- Input validation at boundaries (Pydantic recommended)
- Deployment configuration (Docker, Streamlit Cloud, etc.)

**Important:**
- DO keep calculation logic separate from UI code
- DO validate all external data inputs with Pydantic
- DO use pytest with tests.md for calculation verification
- DO NOT mix data fetching into calculation functions
```

#### Path B Preamble (React + TypeScript)

Include in all instruction files:

```markdown
**What you're receiving:**
- Finished UI designs (React components with full styling)
- Data model definitions (TypeScript types and sample data)
- Test-writing instructions for TDD approach

**What you need to build:**
- Backend API endpoints and database schema
- Authentication and authorization
- Data fetching and state management

**Important:**
- DO NOT redesign the components — use them as-is
- DO wire up callbacks to your routing and APIs
- DO use test-driven development with tests.md
```

### 4. Prepare Files for Export

#### Path A (Python + Streamlit)

- Copy `.py` files with clean imports (no DRIVER-specific paths)
- Generate `requirements.txt` from imports used in the project (pin versions)
- Ensure `calculations/` modules are pure Python (no Streamlit imports)
- Include sample data as CSV/JSON if applicable

#### Path B (React + TypeScript)

When copying components:

- Transform `@/...` to relative paths
- Transform `@/../[project]/build/[section-id]/types` to `../types`
- Remove DRIVER-specific imports

### 5. Create Zip File

After generating all files:

```bash
rm -f driver-plan.zip
zip -r driver-plan.zip driver-plan/
```

### 6. Optimize + Expand (Before Closing)

Before declaring the export complete, the philosophy demands two moves:

**Vertical (Optimize):** What worked well in this project?
- Which prompts or approaches produced the best results? Save them.
- Which sections were built fastest? Why?
- Any reusable patterns worth extracting?

**Horizontal (Expand):** What else does this enable?
- What new questions did this analysis raise?
- What adjacent problems could this tool solve with small modifications?
- What patterns here might transfer to other projects?

Present briefly to the developer:

"Before we wrap up, two quick questions:
1. **What worked best** in this project that you'd want to reuse?
2. **What else could this enable** — any adjacent problems or follow-up questions this tool raises?"

Capture their answers in the export's `README.md` under a "Future Directions" section.

### 7. Confirm Completion

"I've created the complete export package at `driver-plan/` and `driver-plan.zip`.

**What's Included:**

**Prompts:**
- `prompts/one-shot-prompt.md` — Prompt for full implementation
- `prompts/section-prompt.md` — Template for section-by-section

**Instructions:**
- `product-overview.md` — Always provide with any instruction file
- `instructions/one-shot-instructions.md` — All milestones combined
- `instructions/incremental/` — [N] milestone instructions

**Design Assets:**
- `design-system/` — Colors, fonts, tokens
- `data-model/` — Entity types and sample data
- `shell/` — Application shell components
- `sections/` — [N] section component packages with test instructions

**How to Use:**

1. Copy `driver-plan/` to your implementation codebase
2. Open `prompts/one-shot-prompt.md` or `prompts/section-prompt.md`
3. Copy/paste into your AI partner
4. Answer the clarifying questions
5. Let your AI partner implement based on the instructions

**Download:** Restart your dev server and visit the Export page to download `driver-plan.zip`.

---

**This is your final deliverable.** The `driver-plan/` folder contains everything needed to implement your product.

Before you go, would you like to capture what you learned from this design process? It only takes a few minutes and helps improve future projects."

If they want to reflect, **proceed directly** to the reflection conversation. If they're done, wish them well.

---

## Proactive Flow

As a Cognition Mate:
- Generate the complete export automatically
- Suggest reflecting on learnings (optional but valuable)
- If they agree, start the reflection conversation directly

---

## Guiding Principles

- **Final deliverable** — This is what the developer takes away
- **Self-contained** — No dependencies on DRIVER
- **Prompts ask questions** — About auth, data modeling, tech stack
- **TDD support** — Each section has test instructions
- **Show don't tell** — Screenshots provide visual reference
