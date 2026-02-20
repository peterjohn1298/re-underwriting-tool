"""Side-by-side deal comparison."""

import sys
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent.parent))

st.set_page_config(
    page_title="Compare Deals — RE Underwriting",
    page_icon="🔀",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background-color: #0D1117; }
[data-testid="stHeader"] { background-color: #0D1117; }
</style>
""", unsafe_allow_html=True)

st.title("🔀 Deal Comparison")
st.caption("Side-by-side analysis of saved deals.")

# ── Guard ─────────────────────────────────────────────────────────────────────
saved = st.session_state.get("saved_deals", {})
if not saved:
    st.info("No deals saved for comparison yet. Run an analysis and click **📌 Save for Comparison** on the Results page.")
    if st.button("← Back to Input Form"):
        st.switch_page("streamlit_app.py")
    st.stop()


# ── Build comparison table ────────────────────────────────────────────────────
import pandas as pd

deals = list(saved.values())

def _pct(v):
    return f"{v*100:.1f}%" if v is not None else "—"

def _dollar(v):
    return f"${v:,.0f}" if v is not None else "—"

def _x(v):
    return f"{v:.2f}x" if v is not None else "—"

def _fmt(v, d=2):
    return f"{v:.{d}f}" if v is not None else "—"


# Summary comparison table
rows = []
for d in deals:
    rows.append({
        "Property":        d.get("property_name", "—"),
        "Address":         d.get("address", "—"),
        "Purchase Price":  _dollar(d.get("purchase_price")),
        "Units":           str(d.get("total_units", "—")),
        "Levered IRR":     _pct(d.get("levered_irr")),
        "Equity Multiple": _x(d.get("equity_multiple")),
        "Cash-on-Cash Y1": _pct(d.get("cash_on_cash_yr1")),
        "DSCR Y1":         _fmt(d.get("dscr_yr1")),
        "Going-In Cap":    _pct(d.get("going_in_cap_rate")),
        "Exit Cap Rate":   _pct(d.get("exit_cap_rate")),
    })

df = pd.DataFrame(rows)

# Highlight best values
def _highlight_best(col):
    """Green for best IRR/EM/CoC; no color for others."""
    style = [""] * len(col)
    numeric_cols = {"Levered IRR", "Equity Multiple", "Cash-on-Cash Y1"}
    if col.name in numeric_cols:
        # Parse numeric values for comparison
        try:
            vals = []
            for v in col:
                v_str = v.replace("%", "").replace("x", "").replace("$", "").replace(",", "")
                vals.append(float(v_str) if v_str and v_str != "—" else None)
            if any(v is not None for v in vals):
                max_val = max(v for v in vals if v is not None)
                for i, v in enumerate(vals):
                    if v == max_val:
                        style[i] = "background-color: #1a7f37; color: white;"
        except Exception:
            pass
    return style

styled = df.style.apply(_highlight_best, axis=0)
st.dataframe(styled, use_container_width=True, hide_index=True)

# ── NOI Growth Chart ──────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("NOI Growth Over Hold Period")

try:
    import plotly.graph_objects as go
    colors = ["#C9A84C", "#3fb950", "#f85149", "#58a6ff", "#d2a8ff"]
    fig = go.Figure()
    has_data = False
    for i, d in enumerate(deals):
        nois = d.get("annual_nois", [])
        if nois:
            has_data = True
            fig.add_trace(go.Scatter(
                x=list(range(1, len(nois) + 1)),
                y=nois,
                mode="lines+markers",
                name=d.get("property_name", f"Deal {i+1}"),
                line=dict(color=colors[i % len(colors)], width=2),
            ))
    if has_data:
        fig.update_layout(
            xaxis_title="Year",
            yaxis_title="NOI ($)",
            paper_bgcolor="#0D1117",
            plot_bgcolor="#161B22",
            font_color="#E6EDF3",
            legend=dict(bgcolor="#161B22", bordercolor="#30363D"),
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.caption("No NOI data available for charting.")
except ImportError:
    st.caption("Install plotly for NOI growth chart.")

# ── Manage saved deals ────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("Manage Saved Deals")
to_remove = st.multiselect(
    "Remove deals from comparison:",
    options=[d.get("property_name", d.get("job_id", "")) for d in deals],
)
if to_remove and st.button("🗑️ Remove Selected"):
    keys_to_remove = [
        job_id for job_id, d in saved.items()
        if d.get("property_name", d.get("job_id", "")) in to_remove
    ]
    for k in keys_to_remove:
        del st.session_state["saved_deals"][k]
    st.success(f"Removed {len(keys_to_remove)} deal(s).")
    st.rerun()
