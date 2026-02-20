# Validation — Section 3: AI Deal Chat

## Deployment Evidence
- File: pages/2_Chat.py
- Accessible via "💬 AI Chat" button on Results page

## Verified Behaviors
- [x] Guard: redirects to input form if no results in session_state
- [x] API key: reads ANTHROPIC_API_KEY from st.secrets (Streamlit Cloud)
- [x] Warning shown if key not configured
- [x] System prompt built from actual session results (metrics, market, ML, MC, rec, lease)
- [x] ML keys in system prompt use correct keys: assessment, predicted_total_value, actual_price_per_unit
- [x] MC keys in system prompt use correct keys: mean_irr, percentiles["P10/P50/P90"]
- [x] Chat history stored in st.session_state["chat_history"] (last 10 turns)
- [x] Claude claude-opus-4-6 responds grounded in deal data
- [x] "🗑️ Clear Chat" button clears history and reruns
- [x] Model: claude-opus-4-6 (Anthropic) — switched from OpenAI GPT-4o

## Screenshot Instructions (Manual)
Screenshots to capture:
- Chat interface with example question and Claude response
- System prompt showing deal data (debug mode)
