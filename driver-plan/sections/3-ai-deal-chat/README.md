# Section 3: AI Deal Chat

## What It Does
Claude claude-opus-4-6 powered chat assistant grounded entirely in the current deal's analysis
results — cannot hallucinate numbers because the system prompt contains all actual outputs.

## Key Files
- `pages/2_Chat.py` — Chat UI and API integration

## System Prompt Contents
- Deal overview (address, price, units, rent, occupancy, hold period, LTV, rate)
- All key financial metrics (levered/unlevered/after-tax IRR, EM, CoC, DSCR, YOC, cap rates)
- Exit/reversion details (exit year, forward NOI, sale price, net proceeds)
- Pro forma summary (first 3 years: EGI, expenses, NOI, BTCF)
- Market data highlights (demographics, cap rates, crime, walk score, flood zone, HUD FMR, RentCast AVM)
- ML valuation (assessment, predicted total value, actual price/unit, premium/discount)
- Monte Carlo results (P10/P50/P90 IRR, mean, probability table)
- Integrated recommendation (signal, score, all signal breakdown lines)
- Lease analysis summary (lease count, weighted avg escalation)
- Instructions: answer only from data above, flag uncertainty, be concise

## API
- Model: `claude-opus-4-6` (Anthropic)
- Key: `ANTHROPIC_API_KEY` from Streamlit secrets
- Max tokens: 1,024
- History: last 10 turns kept in `st.session_state["chat_history"]`

## Status: COMPLETE
