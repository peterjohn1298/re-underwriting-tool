# RE Underwriting Intelligence Platform — Export Package

This package contains everything needed to understand, extend, or redeploy the
RE Underwriting Intelligence Platform.

## Quick Start

1. Clone the repo: `git clone https://github.com/peterjohn1298/re-underwriting-tool`
2. Install dependencies: `pip install -r requirements.txt`
3. Add API keys to `.streamlit/secrets.toml` (see `secrets.toml.example`)
4. Run locally: `streamlit run streamlit_app.py`
5. Or visit the live deployment: https://re-underwriting-tool.streamlit.app

## What's in This Package

| File/Folder | Contents |
|-------------|----------|
| `product-overview.md` | Problem, success vision, tech stack |
| `prompts/one-shot-prompt.md` | Prompt to rebuild/extend the full tool |
| `prompts/section-prompt.md` | Prompt template for extending one section |
| `instructions/one-shot-instructions.md` | Full implementation guide |
| `instructions/incremental/` | Per-section implementation milestones |
| `sections/*/README.md` | Per-section docs, inputs/outputs, key files |

## Architecture at a Glance

```
streamlit_app.py          ← Entry: input form + full pipeline
pages/
  1_Results.py            ← 11-tab results dashboard
  2_Chat.py               ← Claude AI deal chat
  3_Compare.py            ← Side-by-side deal comparison
models/                   ← Financial models (zero UI dependency)
services/                 ← API clients + AI services
.streamlit/
  secrets.toml            ← API keys (gitignored)
  config.toml             ← Dark theme
```

## API Keys Required

| Secret Key | Service | Used For |
|------------|---------|---------|
| `FRED_API_KEY` | Federal Reserve | Macro rates, CPI, treasury yields |
| `RENTCAST_API_KEY` | RentCast | Rent comps, AVM, market stats |
| `HUD_API_KEY` | HUD | Fair Market Rents by bedroom |
| `GOOGLE_API_KEY` | Google / Gemini | Lease PDF analysis |
| `FBI_CRIME_API_KEY` | FBI UCR | Crime risk data |
| `OPENAI_API_KEY` | OpenAI | AI investment memo (GPT-4o) |
| `ANTHROPIC_API_KEY` | Anthropic | AI deal chat (Claude) |
