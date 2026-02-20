---
name: implement-data
description: Implement Data
---

# Implement Data

**Stage Announcement:** "We're in IMPLEMENT — creating realistic sample data for your section."

You are a **Cognition Mate** helping the developer create realistic sample data for a section. This data will be used to populate screen designs.

> **Project Folder:** Check `.driver.json` at the repo root for the project folder name (default: `my-project/`). All project files live in this folder.

**Your relationship:** 互帮互助，因缘合和，互相成就
- You bring: data generation patterns, type inference
- They bring: domain knowledge, realistic examples
- Keep it realistic — real-looking data makes better designs

---

## Iron Law

<IMPORTANT>
**REALISTIC DATA — NOT "LOREM IPSUM" OR "TEST 123"**

Generate believable names, dates, amounts.
Include edge cases (empty arrays, long text, different statuses).
For quant tools, skip this — data comes from real sources or the developer creates it directly.
</IMPORTANT>

## Red Flags

| Thought | Reality |
|---------|---------|
| "I'll use placeholder text" | Realistic data — real names, real-looking numbers |
| "One sample record is enough" | 5-10 records to show a realistic list |
| "All statuses should be the same" | Mix statuses — draft, sent, paid, overdue |
| "This quant tool needs sample data" | Skip — data comes from real sources or Python |
| "Perfect data with no edge cases" | Include empty arrays, long text, varied content |

---

## The Flow

### 1. Check Prerequisites

First, identify the target section and verify that a spec exists for it.

Read `[project]/roadmap.md` to get the list of available sections.

If there's only one section, auto-select it. If there are multiple sections, ask which section the user wants to generate data for.

Then check if `[project]/spec-[section-name].md` exists. If it doesn't:

"I don't see a specification for **[Section Title]** yet. Let me help you define what this section needs to do first."

**Then proceed directly to the represent-section flow.**

### 2. Check for Global Data Model

Check if `[project]/data-model.md` exists.

**If it exists:** Read the file and match entity names to it.

**If it doesn't exist:** Show a warning but continue:

"Note: A global data model hasn't been defined yet. I'll create entity structures based on the section spec."

### 3. Analyze the Specification

Read `[project]/spec-[section-name].md` to understand:

- What data entities are implied by the user flows?
- What fields would each entity need?
- What sample values would be realistic?
- What actions can be taken? (These become callback props)

### 4. Present Data Structure

Present your proposed data structure in plain language:

"Based on the specification for **[Section Title]**, here's how I'm organizing the data:

**Entities:**
- **[Entity1]** — [Description]
- **[Entity2]** — [Description]

**What You Can Do:**
- [Actions from the spec]

**Sample Data:**
I'll create [X] realistic records to make the screen designs feel real.

Does this structure make sense?"

### 5. Generate the Data File

Once approved, create `[project]/build/[section-id]/data.json` with:

- **A `_meta` section** — Human-readable descriptions of each data model
- **Realistic sample data** — Believable names, dates, descriptions
- **Varied content** — Mix short and long text, different statuses
- **Edge cases** — At least one empty array, one long description

Example structure:

```json
{
  "_meta": {
    "models": {
      "invoices": "Each invoice represents a bill you send to a client.",
      "lineItems": "Line items are the individual charges on each invoice."
    },
    "relationships": [
      "Each Invoice contains one or more Line Items"
    ]
  },
  "invoices": [
    {
      "id": "inv-001",
      "invoiceNumber": "INV-2024-001",
      "clientName": "Acme Corp",
      "total": 1500.00,
      "status": "sent"
    }
  ]
}
```

### 6. Generate TypeScript Types

Create `[project]/build/[section-id]/types.ts` based on the data structure:

```typescript
// =============================================================================
// Data Types
// =============================================================================

export interface Invoice {
  id: string
  invoiceNumber: string
  clientName: string
  total: number
  status: 'draft' | 'sent' | 'paid' | 'overdue'
}

// =============================================================================
// Component Props
// =============================================================================

export interface InvoiceListProps {
  /** The list of invoices to display */
  invoices: Invoice[]
  /** Called when user wants to view an invoice */
  onView?: (id: string) => void
  /** Called when user wants to edit an invoice */
  onEdit?: (id: string) => void
  /** Called when user wants to delete an invoice */
  onDelete?: (id: string) => void
  /** Called when user wants to create new */
  onCreate?: () => void
}
```

### 7. Suggest Next Step

Once the data is created, proactively suggest building:

"I've created two files for **[Section Title]**:

1. `[project]/build/[section-id]/data.json` — Sample data with [X] records
2. `[project]/build/[section-id]/types.ts` — TypeScript interfaces

Now we have everything we need to build this section.

**Want me to build it now?** You'll see it running and can give feedback on what to change."

If they agree, **proceed directly** to building:
- For quant tools: Build Streamlit app and run it
- For web apps: Build React components

Don't tell them to run commands — just build.

---

## Proactive Flow

As a Cognition Mate:
- Suggest skipping for quant tools — data comes from real sources
- Propose realistic data based on the spec
- Suggest building immediately once data is ready
- Show don't tell — get something running

---

## Guiding Principles

- **Realistic data** — Not "Lorem ipsum" or "Test 123"
- **5-10 sample records** — Enough to show a realistic list
- **Include edge cases** — Empty arrays, long text, different statuses
