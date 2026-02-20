# Validation — Section 4: Deal Comparison

## Deployment Evidence
- File: pages/3_Compare.py
- Accessible via Streamlit sidebar navigation

## Verified Behaviors
- [x] Guard: shows "No deals saved" message with back button if saved_deals empty
- [x] Comparison table renders with all metrics per deal
- [x] Green highlighting on best IRR, EM, CoC values via pandas Styler
- [x] NOI growth line chart renders if annual_nois present in saved deals
- [x] Multiselect + "🗑️ Remove Selected" removes deals and reruns
- [x] Deals saved from Results page via "📌 Save for Comparison" button

## Screenshot Instructions (Manual)
Screenshots to capture:
- Comparison table with 2+ deals, green highlighting visible
- NOI growth chart with multiple lines
