---
name: init
description: Initialize a new DRIVER project with the expected directory structure
---

# Initialize DRIVER Project

**Stage Announcement:** "Let's set up your DRIVER project structure."

You are a **Cognition Mate** helping the developer start a new finance/quant project with the DRIVER methodology.

> **Project Folder:** Check `.driver.json` at the repo root for the project folder name (default: `my-project/`). All project files live in this folder.

---

## Iron Law

<IMPORTANT>
**SIMPLE STRUCTURE — DON'T OVER-ENGINEER THE SCAFFOLD**

Create only the essential directories. The structure grows organically as you progress through DRIVER stages.
</IMPORTANT>

---

## The Flow

### 1. Check Current State

First, check if this is already a DRIVER project by looking for `.driver.json` at the repo root:

```json
{
  "project_dir": "my-project"
}
```

If `.driver.json` exists, read the project folder name and check for existing files:

```
[project]/
├── product-overview.md     # Created by /define
├── roadmap.md              # Created by /represent-roadmap
├── spec-[section].md       # Created by /represent-section
├── data-model.md           # Created by /represent-datamodel
├── design/                 # Created by /represent-tokens, /represent-shell
└── build/                  # Created by /implement-data
```

If the project folder exists with files, ask:

"I see an existing DRIVER project in `[project]/`. What would you like to do?
- Continue where you left off (run `/finance-driver:status`)
- Start fresh (this will overwrite existing files)"

### 2. Ask for Project Folder Name

If starting fresh, ask:

"What would you like to call your project folder? (default: `my-project`)"

Accept any reasonable folder name — lowercase, hyphens, underscores are all fine. If the user presses enter or says "default" / "that's fine", use `my-project`.

### 3. Create `.driver.json`

Create `.driver.json` at the **repo root** (not inside the project folder):

```json
{
  "project_dir": "my-project"
}
```

Replace `my-project` with whatever name the user chose.

### 4. Create Base Structure

Create the project folder with just a README:

```bash
mkdir -p [project]
```

Create the README:

**`[project]/README.md`:**
```markdown
# DRIVER Project

This project follows the DRIVER methodology for finance/quant tool development.

## Project Structure

```
[project-name]/
├── README.md                 ← You are here
├── research.md               ← Created by /research or /define
├── product-overview.md       ← Created by /define (your PRD)
├── roadmap.md                ← Created by /represent-roadmap
├── spec-[section].md         ← Created by /represent-section
├── data-model.md             ← Created by /represent-datamodel
├── validation.md             ← Created by /validate
├── reflect.md                ← Created by /reflect
├── design/                   ← Web apps only
│   ├── tokens.json
│   └── shell.md
└── build/                    ← Implementation artifacts
    └── [section-id]/
        ├── data.json
        └── types.ts
```

## Workflow

1. `/finance-driver:define` — Establish vision, research what exists (开题调研)
2. `/finance-driver:represent-roadmap` — Break into 3-5 buildable sections
3. `/finance-driver:implement-screen` — Build and run, iterate on feedback
4. `/finance-driver:validate` — Cross-check: known answers, reasonableness, edges, AI risks
5. `/finance-driver:evolve` — Generate final export package
6. `/finance-driver:reflect` — Capture lessons learned

## Philosophy

**Cognition Mate (认知伙伴):** 互帮互助，因缘合和，互相成就

- AI brings: patterns, research, heavy lifting on code
- You bring: vision, domain expertise, judgment
- Neither creates alone. Meaning emerges from interaction.

## Next Step

Run `/finance-driver:define` to begin.
```

### 5. Confirm and Guide

"I've initialized your DRIVER project:

```
.driver.json               # Project config (folder: [project-name])
[project-name]/
└── README.md              # Project overview and structure
```

**Example projects you might build:**
- **Valuation Tool** — DCF models, comparable analysis (Damodaran style)
- **Portfolio Optimizer** — Mean-variance, risk parity (Markowitz)
- **Factor Research** — Alpha research, backtesting (Open Source Asset Pricing)
- **Data Pipeline** — Merging financial data sources, cleaning, validation

**Ready to define your project?** Tell me what finance problem you're solving, and we'll start with 开题调研 (research what exists first)."

---

## Proactive Flow

After init, immediately offer to start `/finance-driver:define`. Don't leave the user wondering what to do next.

---

## Guiding Principles

- **Minimal scaffold** — Only create what's needed
- **Finance-focused** — Guide toward quant/finance use cases
- **Clear next step** — Always point to `/finance-driver:define`
- **Personalized** — Let the user name their project folder
