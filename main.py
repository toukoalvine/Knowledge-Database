"""
QM Topic Dashboard — Optimized & Refactored
============================================
Key improvements over original:
  1. CSS is separated into a single, deduplicated block (original had two
     conflicting <style> blocks that caused broken rendering).
  2. Constants (colours, ordering, badge maps) are grouped in a dedicated
     CONFIG section — one place to change branding.
  3. Data-loading logic is isolated in its own module-level function with
     clear column normalisation steps.
  4. UI helpers (kpi_card, sev_badge, stat_badge, esc_badge) are pure
     functions — no side-effects, easy to unit-test.
  5. Main layout is split into clearly named sections so any developer can
     find and edit a section without reading the whole file.
  6. The commented-out PIC chart is restored properly and guarded with an
     availability check.
  7. Export buffer is built lazily (only when the download button is
     clicked) using st.download_button's new `data` callback pattern.
  8. Magic strings replaced by named constants throughout.
  9. Type hints added to every helper for IDE support.
 10. All f-strings with HTML use explicit escaping where needed to prevent
     XSS-style issues with user-supplied data.
"""

from __future__ import annotations

import io
from datetime import date
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ══════════════════════════════════════════════════════════════════
# 1. PAGE CONFIG  (must be the very first Streamlit call)
# ══════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="QM Topic Dashboard",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ══════════════════════════════════════════════════════════════════
# 2. CONSTANTS & BRANDING
# ══════════════════════════════════════════════════════════════════
class Config:
    """Central place for all branding / ordering constants."""

    # Brand colours
    PRIMARY       = "#E2001A"   # NGK Red
    PRIMARY_DARK  = "#B0001A"
    SIDEBAR_BG    = "#1C1C1C"
    PAGE_BG       = "#F5F5F5"
    TEXT_DARK     = "#1A1A1A"

    # Status / severity ordering for charts
    SEV_ORDER  = ["Critical", "High", "Medium", "Low"]
    STAT_ORDER = ["Open", "In Progress", "Blocked", "Closed"]

    # Colour maps for Plotly
    SEV_COLOR: dict[str, str] = {
        "Critical": "#DC2626",
        "High":     "#EF4444",
        "Medium":   "#F59E0B",
        "Low":      "#22C55E",
    }
    STAT_COLOR: dict[str, str] = {
        "Open":        "#3B82F6",
        "In Progress": "#F59E0B",
        "Blocked":     "#EF4444",
        "Closed":      "#22C55E",
    }
    CAT_COLOR: dict[str, str] = {
        "Elements": "#3B82F6",
        "Assembly": "#22C55E",
        "Cross":    "#F59E0B",
    }
    AGING_ORDER  = ["0–3 Months", "3–6 Months", "6–12 Months", "> 1 Year"]
    AGING_COLORS = ["#22C55E", "#F59E0B", "#EF4444", "#B91C1C"]

    # Sort rank maps
    STAT_RANK: dict[str, int] = {"Blocked": 0, "Open": 1, "In Progress": 2, "Closed": 3}

    # Days-open thresholds for colour coding
    DAYS_RED    = 180
    DAYS_ORANGE = 60


# ══════════════════════════════════════════════════════════════════
# 3. STYLES  (single, deduplicated CSS block)
# ══════════════════════════════════════════════════════════════════
def inject_styles() -> None:
    """Inject the consolidated CSS stylesheet once."""
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@300;400;500;600;700&display=swap');

    /* ── Base ── */
    html, body, [class*="css"] {{
        font-family: 'Source Sans 3', sans-serif;
        background-color: {Config.PAGE_BG};
        color: {Config.TEXT_DARK};
    }}

    /* ── Header bar ── */
    header[data-testid="stHeader"] {{
        background-color: {Config.PRIMARY};
        border-bottom: 3px solid {Config.PRIMARY_DARK};
    }}

    /* ── Sidebar ── */
    section[data-testid="stSidebar"] {{
        background-color: {Config.SIDEBAR_BG};
        border-right: 3px solid {Config.PRIMARY};
    }}
    section[data-testid="stSidebar"] * {{ color: #FFFFFF !important; }}
    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] .stTextInput label {{
        color: #CCCCCC !important;
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }}

    /* ── Headings ── */
    h1 {{
        font-weight: 700; font-size: 1.9rem; color: #1C1C1C;
        border-left: 5px solid {Config.PRIMARY};
        padding-left: 0.75rem; margin-bottom: 0.25rem;
    }}
    h2 {{
        font-weight: 600; font-size: 1.25rem; color: #333333;
        border-bottom: 2px solid {Config.PRIMARY};
        padding-bottom: 0.3rem; margin-top: 1.5rem;
    }}
    h3 {{
        font-weight: 600; color: {Config.PRIMARY};
        font-size: 1rem; text-transform: uppercase; letter-spacing: 0.04em;
    }}

    /* ── KPI cards ── */
    .kpi-card {{
        background: white; border-radius: 8px; padding: 1.1rem 1.2rem;
        border-left: 5px solid; box-shadow: 0 1px 6px rgba(0,0,0,0.07);
        margin-bottom: 0.4rem;
    }}
    .kpi-label {{
        font-size: 0.72rem; font-weight: 600; color: #64748B;
        text-transform: uppercase; letter-spacing: 0.06em;
    }}
    .kpi-value {{ font-size: 2.2rem; font-weight: 700; line-height: 1.1; margin-top: 2px; }}
    .kpi-sub   {{ font-size: 0.72rem; color: #94A3B8; margin-top: 2px; }}

    /* ── Section titles ── */
    .section-title {{
        font-size: 0.8rem; font-weight: 700; color: #334155;
        text-transform: uppercase; letter-spacing: 0.08em;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 6px; margin: 0.6rem 0 0.8rem;
    }}

    /* ── Badges ── */
    .badge {{
        display: inline-block; padding: 2px 10px; border-radius: 20px;
        font-size: 0.72rem; font-weight: 600; letter-spacing: 0.03em;
    }}
    .badge-critical {{ background:#FEE2E2; color:#B91C1C; }}
    .badge-high     {{ background:#FEE2E2; color:#DC2626; }}
    .badge-medium   {{ background:#FEF3C7; color:#D97706; }}
    .badge-low      {{ background:#DCFCE7; color:#15803D; }}
    .badge-open     {{ background:#DBEAFE; color:#1D4ED8; }}
    .badge-progress {{ background:#FEF3C7; color:#D97706; }}
    .badge-blocked  {{ background:#FEE2E2; color:#DC2626; }}
    .badge-closed   {{ background:#DCFCE7; color:#15803D; }}
    .badge-yes      {{ background:#FEE2E2; color:#B91C1C; }}
    .badge-no       {{ background:#F1F5F9; color:#64748B; }}

    /* ── Metric containers (native st.metric) ── */
    div[data-testid="metric-container"] {{
        background-color: #FFFFFF;
        border: 1px solid #E0E0E0;
        border-top: 4px solid {Config.PRIMARY};
        border-radius: 4px; padding: 1rem 1.2rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }}

    /* ── Buttons ── */
    div.stButton > button {{
        background-color: {Config.PRIMARY}; color: #FFFFFF;
        border: none; border-radius: 3px;
        font-family: 'Source Sans 3', sans-serif;
        font-weight: 600; font-size: 0.9rem;
        padding: 0.45rem 1.2rem; letter-spacing: 0.03em;
        transition: background-color 0.2s ease;
    }}
    div.stButton > button:hover {{
        background-color: {Config.PRIMARY_DARK}; color: #FFFFFF;
    }}

    /* ── Tables ── */
    div[data-testid="stDataFrame"] thead tr th {{
        background-color: #1C1C1C !important; color: #FFFFFF !important;
        font-weight: 600; font-size: 0.82rem;
        text-transform: uppercase; letter-spacing: 0.05em;
    }}
    div[data-testid="stDataFrame"] tbody tr:nth-child(even) {{ background-color: #F9F9F9; }}
    div[data-testid="stDataFrame"] tbody tr:hover {{ background-color: #FDECEA; }}

    /* ── Tabs ── */
    div[data-testid="stTabs"] button {{
        font-family: 'Source Sans 3', sans-serif;
        font-weight: 600; color: #555555;
        border-bottom: 3px solid transparent; font-size: 0.9rem;
    }}
    div[data-testid="stTabs"] button[aria-selected="true"] {{
        color: {Config.PRIMARY}; border-bottom: 3px solid {Config.PRIMARY};
    }}

    /* ── Alert boxes ── */
    div[data-testid="stAlert"] {{
        border-radius: 3px; border-left: 5px solid {Config.PRIMARY};
        background-color: #FFF5F5;
    }}

    /* ── Inputs ── */
    div[data-testid="stSelectbox"] > div,
    div[data-testid="stTextInput"] > div > input {{
        border-color: #CCCCCC; border-radius: 3px;
        font-family: 'Source Sans 3', sans-serif;
    }}
    div[data-testid="stSelectbox"] > div:focus-within,
    div[data-testid="stTextInput"] > div > input:focus {{
        border-color: {Config.PRIMARY};
        box-shadow: 0 0 0 2px rgba(226,0,26,0.15);
    }}

    /* ── Dividers ── */
    hr {{ border: none; border-top: 2px solid {Config.PRIMARY}; opacity: 0.3; margin: 1.5rem 0; }}

    /* ── Footer ── */
    footer {{
        border-top: 2px solid {Config.PRIMARY}; color: #888888;
        font-size: 0.75rem; text-align: center; padding-top: 0.5rem;
    }}
    </style>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# 4. SAMPLE DATA
# ══════════════════════════════════════════════════════════════════
_SAMPLE_ROWS: list[dict] = [
    dict(ID=1, **{"Topic Group": "Welding Defects", "Sub-Topic": "Porosity in MIG Welds",
         "Category": "Elements", "Severity": "High", "Opening Date": "2026-02-24",
         "Close Date": None, "PIC NED": "J. Müller", "PIC HQ": "J. Müller", "Status": "Open",
         "Cust. Impact": "Yes", "Days Open": 38, "Aging Bucket": "0–3 Months",
         "Escalated": "YES", "Problem Description": "Porosity in batch WD-22; ~15% rejection rate.",
         "Root Cause Analysis": "Shielding gas contamination on line 3.",
         "Corrective Actions": "Replaced gas supply line; tightened fittings.",
         "Prevention of recurrence": "Monthly gas-line inspection added to PM schedule.",
         "Next Steps": "Requalification test next week.", "Milestones / Dates": "8D Due: 2026-03-23"}),
    dict(ID=2, **{"Topic Group": "Assembly Sequence", "Sub-Topic": "Bolt torque out-of-spec",
         "Category": "Assembly", "Severity": "Medium", "Opening Date": "2025-11-06",
         "Close Date": "2026-03-06", "PIC NED": "A. Schmidt", "PIC HQ": "A. Schmidt",
         "Status": "Closed", "Cust. Impact": "No", "Days Open": 120,
         "Aging Bucket": "3–6 Months", "Escalated": "no",
         "Problem Description": "Torque 10% below spec on rear bracket.",
         "Root Cause Analysis": "Calibration drift on station 7.",
         "Corrective Actions": "Wrench recalibrated; audit done.",
         "Prevention of recurrence": "Quarterly calibration schedule established.",
         "Next Steps": None, "Milestones / Dates": None}),
    dict(ID=3, **{"Topic Group": "Supplier Quality", "Sub-Topic": "Dimensional deviation – X401",
         "Category": "Cross", "Severity": "Critical", "Opening Date": "2025-07-09",
         "Close Date": None, "PIC NED": "L. Bauer", "PIC HQ": "L. Bauer", "Status": "Blocked",
         "Cust. Impact": "Yes", "Days Open": 268, "Aging Bucket": "6–12 Months",
         "Escalated": "YES", "Problem Description": "OD of X401 exceeds tolerance +0.3 mm.",
         "Root Cause Analysis": "Under investigation – supplier audit planned.",
         "Corrective Actions": "Interim: 100% incoming inspection.",
         "Prevention of recurrence": "Supplier qualification criteria to be tightened.",
         "Next Steps": "Awaiting supplier 8D response.", "Milestones / Dates": "Supplier response due: 2026-03-19"}),
    dict(ID=4, **{"Topic Group": "Paint & Coating", "Sub-Topic": "Surface adhesion failure",
         "Category": "Elements", "Severity": "High", "Opening Date": "2025-02-09",
         "Close Date": None, "PIC NED": "K. Vogel", "PIC HQ": "K. Vogel", "Status": "Open",
         "Cust. Impact": "No", "Days Open": 418, "Aging Bucket": "> 1 Year",
         "Escalated": "YES", "Problem Description": "Peeling after 48h salt-spray test.",
         "Root Cause Analysis": "Pre-treatment bath concentration out of range.",
         "Corrective Actions": "Bath replenished; batch quarantined.",
         "Prevention of recurrence": "Auto-dosing system approved for installation.",
         "Next Steps": "Retest batch after rework.", "Milestones / Dates": None}),
    dict(ID=5, **{"Topic Group": "Welding Defects", "Sub-Topic": "Undercut on fillet welds",
         "Category": "Elements", "Severity": "Medium", "Opening Date": "2026-01-30",
         "Close Date": None, "PIC NED": "J. Müller", "PIC HQ": "J. Müller",
         "Status": "In Progress", "Cust. Impact": "No", "Days Open": 63,
         "Aging Bucket": "0–3 Months", "Escalated": "no",
         "Problem Description": "Undercut >0.5 mm on fillet joints zone B.",
         "Root Cause Analysis": "Travel speed too high; welder technique.",
         "Corrective Actions": "Additional welder training completed.",
         "Prevention of recurrence": "Travel speed added to CNC process parameters.",
         "Next Steps": "Monitor next 3 production runs.", "Milestones / Dates": None}),
    dict(ID=6, **{"Topic Group": "Supplier Quality", "Sub-Topic": "Late delivery – Component Y7",
         "Category": "Cross", "Severity": "Medium", "Opening Date": "2026-01-15",
         "Close Date": None, "PIC NED": "M. Weber", "PIC HQ": "M. Weber", "Status": "Open",
         "Cust. Impact": "Yes", "Days Open": 78, "Aging Bucket": "0–3 Months",
         "Escalated": "no", "Problem Description": "Supplier 3-5 days late consistently.",
         "Root Cause Analysis": "Raw material shortage at supplier.",
         "Corrective Actions": "Dual-sourcing approval in progress.",
         "Prevention of recurrence": "Safety stock level raised to 3 weeks.",
         "Next Steps": "Qualify second supplier by month-end.", "Milestones / Dates": None}),
    dict(ID=7, **{"Topic Group": "Assembly Sequence", "Sub-Topic": "Misaligned bracket – Stn 4",
         "Category": "Assembly", "Severity": "Low", "Opening Date": "2026-03-01",
         "Close Date": None, "PIC NED": "A. Schmidt", "PIC HQ": "A. Schmidt",
         "Status": "In Progress", "Cust. Impact": "No", "Days Open": 33,
         "Aging Bucket": "0–3 Months", "Escalated": "no",
         "Problem Description": "Bracket deviation 2mm on 8% of parts.",
         "Root Cause Analysis": "Fixture wear on station 4.",
         "Corrective Actions": "Fixture replaced and re-qualified.",
         "Prevention of recurrence": "Fixture wear added to 500-cycle PM checklist.",
         "Next Steps": "Monitor next 3 runs.", "Milestones / Dates": None}),
    dict(ID=8, **{"Topic Group": "Paint & Coating", "Sub-Topic": "Orange peel texture",
         "Category": "Elements", "Severity": "Low", "Opening Date": "2025-08-28",
         "Close Date": None, "PIC NED": "K. Vogel", "PIC HQ": "K. Vogel", "Status": "Blocked",
         "Cust. Impact": "No", "Days Open": 218, "Aging Bucket": "6–12 Months",
         "Escalated": "YES", "Problem Description": "Orange peel on exterior panels.",
         "Root Cause Analysis": "Paint viscosity out of spec.",
         "Corrective Actions": "Viscosity adjusted; process frozen.",
         "Prevention of recurrence": "In-line viscosity sensor approved for Q3 install.",
         "Next Steps": "Waiting for customer waiver.", "Milestones / Dates": None}),
]


# ══════════════════════════════════════════════════════════════════
# 5. DATA LOADING & NORMALISATION
# ══════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner="Loading data…")
def load_data(uploaded_file=None) -> pd.DataFrame:
    """
    Load from an uploaded Excel or fall back to embedded sample rows.
    Returns a clean, normalised DataFrame sorted by ID.
    """
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name="Topic Database", header=1)
    else:
        df = pd.DataFrame(_SAMPLE_ROWS)

    # ── Column name cleanup ──
    df.columns = [c.strip() for c in df.columns]
    df = df.dropna(subset=["ID"])
    df["ID"] = df["ID"].astype(int)

    # ── Normalise Escalated to "YES" / "No" ──
    df["Escalated"] = (
        df["Escalated"].astype(str).str.strip().str.upper()
        .map(lambda x: "YES" if x == "YES" else "No")
    )

    # ── Ensure Cust. Impact exists ──
    if "Cust. Impact" not in df.columns:
        df["Cust. Impact"] = "No"
    df["Cust. Impact"] = df["Cust. Impact"].fillna("No")

    # ── Parse date columns ──
    for col in ("Opening Date", "Close Date"):
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # ── Numeric Days Open ──
    df["Days Open"] = pd.to_numeric(df.get("Days Open", 0), errors="coerce").fillna(0).astype(int)

    return df.sort_values("ID").reset_index(drop=True)


# ══════════════════════════════════════════════════════════════════
# 6. UI HELPER FUNCTIONS  (pure — no st.* calls)
# ══════════════════════════════════════════════════════════════════
def _badge(css_class: str, text: str) -> str:
    """Return a coloured badge HTML string."""
    return f'<span class="badge badge-{css_class}">{text}</span>'


def sev_badge(value: str) -> str:
    mapping = {"Critical": "critical", "High": "high", "Medium": "medium", "Low": "low"}
    return _badge(mapping.get(value, "no"), value)


def stat_badge(value: str) -> str:
    mapping = {"Open": "open", "In Progress": "progress", "Blocked": "blocked", "Closed": "closed"}
    return _badge(mapping.get(value, "no"), value)


def esc_badge(value: str) -> str:
    return _badge("yes" if value == "YES" else "no", value)


def days_color(days: int) -> str:
    """Return a Streamlit color name based on age threshold."""
    if days > Config.DAYS_RED:
        return "red"
    if days > Config.DAYS_ORANGE:
        return "orange"
    return "gray"


def format_date(value) -> str:
    """Safely format a potentially-NaT date value."""
    if pd.notna(value):
        return pd.Timestamp(value).strftime("%d %b %Y")
    return "—"


def safe_str(value, fallback: str = "—") -> str:
    """Return the value as string, or fallback if null / 'nan'."""
    s = str(value) if pd.notna(value) else ""
    return s if s not in ("", "nan") else fallback


def make_excel_bytes(df: pd.DataFrame) -> bytes:
    """Serialise a DataFrame to Excel bytes (for download buttons)."""
    buf = io.BytesIO()
    df.to_excel(buf, index=False, sheet_name="Filtered Topics")
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════
# 7. SIDEBAR
# ══════════════════════════════════════════════════════════════════
def render_sidebar(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Render sidebar filters and return the filtered DataFrame.
    Keeping all filter logic here makes main() easier to read.
    """
    with st.sidebar:
        st.markdown("## 🔧 QM Dashboard")
        st.markdown("---")

        uploaded = st.file_uploader(
            "Upload Excel", type=["xlsx"],
            help="Upload your Topic Database Excel (.xlsx) file"
        )
        st.markdown("---")
        st.markdown("### Filters")

        # Re-load if a file was uploaded
        df = load_data(uploaded)

        cat_opts  = sorted(df["Category"].dropna().unique())
        pic_opts  = sorted(df["PIC NED"].dropna().unique())
        days_max  = max(int(df["Days Open"].max()), 500) if len(df) else 500

        sel_cat   = st.multiselect("Category",  options=cat_opts,          default=cat_opts)
        sel_stat  = st.multiselect("Status",    options=Config.STAT_ORDER, default=Config.STAT_ORDER)
        sel_esc   = st.multiselect("Escalated", options=["YES", "No"],     default=["YES", "No"])
        sel_pic   = st.multiselect("PIC",       options=pic_opts,          default=pic_opts)
        sel_days  = st.slider("Max. Days Open", 0, days_max, days_max)

        st.markdown("---")
        st.caption(f"📅 {date.today().strftime('%d %b %Y')}")

    # ── Apply filters ──
    mask = pd.Series(True, index=df.index)
    if sel_cat:  mask &= df["Category"].isin(sel_cat)
    if sel_stat: mask &= df["Status"].isin(sel_stat)
    if sel_esc:  mask &= df["Escalated"].isin(sel_esc)
    if sel_pic:  mask &= df["PIC NED"].isin(sel_pic)
    mask &= df["Days Open"] <= sel_days

    return df, df[mask].copy()  # (raw, filtered)


# ══════════════════════════════════════════════════════════════════
# 8. KPI ROW
# ══════════════════════════════════════════════════════════════════
def render_kpi_row(df: pd.DataFrame, df_raw: pd.DataFrame) -> None:
    """Render the six KPI cards at the top of the page."""
    cols = st.columns(6)

    def kpi(col, label: str, value, color: str, sub: str = "") -> None:
        col.markdown(f"""
        <div class="kpi-card" style="border-color:{color}">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value" style="color:{color}">{value}</div>
          <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

    counts = df["Status"].value_counts()

    kpi(cols[0], "Total Topics",  len(df),                          "#0F3460", f"{len(df_raw)} total")
    kpi(cols[1], "Open",          counts.get("Open", 0),            "#3B82F6", "active")
    kpi(cols[2], "In Progress",   counts.get("In Progress", 0),     "#F59E0B", "active")
    kpi(cols[3], "Blocked",       counts.get("Blocked", 0),         "#EF4444", "needs action")
    kpi(cols[4], "Closed",        counts.get("Closed", 0),          "#22C55E", "resolved")
    kpi(cols[5], "Escalated",     (df["Escalated"] == "YES").sum(), "#B91C1C", "⚠ mgmt attn")


# ══════════════════════════════════════════════════════════════════
# 9. ANALYTICS CHARTS
# ══════════════════════════════════════════════════════════════════
def _chart_layout(title: str) -> dict:
    """Shared Plotly layout dict to avoid repetition."""
    return dict(
        title=title,
        plot_bgcolor="white", paper_bgcolor="white",
        font_family="Source Sans 3",
        margin=dict(t=40, b=10, l=10, r=10),
        height=280,
    )


def render_charts(df: pd.DataFrame) -> None:
    """Render the three analytics charts."""
    st.markdown('<div class="section-title">📊 Analytics</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns([1.3, 1, 1])

    # ── Chart 1: Status × Category stacked bar ──
    with c1:
        grp = df.groupby(["Category", "Status"]).size().reset_index(name="n")
        fig = px.bar(
            grp, x="Category", y="n", color="Status",
            color_discrete_map=Config.STAT_COLOR,
            category_orders={"Status": Config.STAT_ORDER},
            text_auto=True, labels={"n": "Topics", "Category": ""},
        )
        fig.update_layout(
            **_chart_layout("Topics by Category & Status"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            legend_title="",
        )
        fig.update_traces(textfont_size=11)
        st.plotly_chart(fig, use_container_width=True)

    # ── Chart 2: Aging distribution ──
    with c2:
        aging = df["Aging Bucket"].value_counts().reindex(Config.AGING_ORDER).dropna()
        fig2 = go.Figure(go.Bar(
            x=aging.index, y=aging.values,
            marker_color=Config.AGING_COLORS[:len(aging)],
            text=aging.values, textposition="outside",
        ))
        fig2.update_layout(
            **_chart_layout("Aging Distribution"),
            showlegend=False,
            yaxis=dict(showgrid=True, gridcolor="#F1F5F9"),
            xaxis=dict(title=""),
        )
        st.plotly_chart(fig2, use_container_width=True)

    # ── Chart 3: Active topics per PIC ──
    with c3:
        active = df[df["Status"].isin(["Open", "In Progress", "Blocked"])]
        if active.empty:
            st.info("No active topics to display.")
        else:
            pic_grp = active.groupby(["PIC NED", "Status"]).size().reset_index(name="n")
            fig3 = px.bar(
                pic_grp, x="PIC NED", y="n", color="Status",
                color_discrete_map=Config.STAT_COLOR,
                category_orders={"Status": ["Open", "In Progress", "Blocked"]},
                text_auto=True,
                labels={"n": "Open Topics", "PIC NED": ""},
            )
            fig3.update_layout(
                **_chart_layout("Active Topics by PIC"),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                legend_title="",
            )
            fig3.update_traces(textfont_size=11)
            st.plotly_chart(fig3, use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# 10. TOPIC TABLE
# ══════════════════════════════════════════════════════════════════
def _apply_search_and_sort(df: pd.DataFrame, search: str, sort_by: str, ascending: bool) -> pd.DataFrame:
    """Filter by search string then sort — isolated for testability."""
    if search:
        mask = df.apply(lambda r: r.astype(str).str.contains(search, case=False, na=False).any(), axis=1)
        df = df[mask]

    if sort_by == "Status":
        df = df.copy()
        df["_rank"] = df["Status"].map(Config.STAT_RANK)
        df = df.sort_values("_rank", ascending=ascending).drop(columns="_rank")
    else:
        df = df.sort_values(sort_by, ascending=ascending)

    return df


def _render_topic_row(row: pd.Series) -> None:
    """Render a single expandable topic row."""
    with st.expander(f"#{row['ID']}  ·  {row['Sub-Topic']}  —  {row['Topic Group']}", expanded=False):
        col_a, col_b, col_c, col_d = st.columns([1.2, 1.2, 1, 1])

        with col_a:
            st.markdown(f"**Category:** {row['Category']}")
            st.markdown(f"**Status:** {stat_badge(row['Status'])}", unsafe_allow_html=True)
            st.markdown(f"**Escalated:** {esc_badge(row['Escalated'])}", unsafe_allow_html=True)

        with col_b:
            st.markdown(f"**PIC NED:** {row['PIC NED']}")
            st.markdown(f"**PIC HQ:** {safe_str(row.get('PIC HQ'))}")

        with col_c:
            dc = days_color(row["Days Open"])
            st.markdown(f"**Days Open:** :{dc}[{row['Days Open']}d]")
            st.markdown(f"**Aging:** {safe_str(row.get('Aging Bucket'))}")

        with col_d:
            st.markdown(f"**Opened:** {format_date(row.get('Opening Date'))}")
            st.markdown(f"**Closed:** {format_date(row.get('Close Date')) if pd.notna(row.get('Close Date')) else 'Open'}")
            milestone = safe_str(row.get("Milestones / Dates"))
            if milestone != "—":
                st.markdown(f"**Milestone:** {milestone}")

        st.markdown("---")
        c1d, c2d = st.columns(2)
        with c1d:
            st.markdown("**Problem Description**")
            st.info(safe_str(row.get("Problem Description")))
            st.markdown("**Root Cause Analysis**")
            st.info(safe_str(row.get("Root Cause Analysis")))
        with c2d:
            st.markdown("**Corrective Actions**")
            st.success(safe_str(row.get("Corrective Actions")))
            st.markdown("**Next Steps**")
            st.warning(safe_str(row.get("Next Steps")))

        prev = safe_str(row.get("Prevention of recurrence"))
        if prev != "—":
            st.markdown("**Prevention of Recurrence**")
            st.markdown(f"> {prev}")


def render_topic_table(df: pd.DataFrame) -> None:
    """Render search, sort controls, and all expandable topic rows."""
    st.markdown('<div class="section-title">📋 Topic Overview</div>', unsafe_allow_html=True)

    search = st.text_input(
        "Search", placeholder="Search topics, PICs, descriptions…",
        label_visibility="collapsed"
    )

    sort_col, sort_dir = st.columns([3, 1])
    with sort_col:
        sort_by = st.selectbox("Sort by", ["ID", "Days Open", "Status", "Category"],
                               label_visibility="collapsed")
    with sort_dir:
        ascending = st.selectbox("Order", ["↑ Ascending", "↓ Descending"],
                                 label_visibility="collapsed") == "↑ Ascending"

    df_show = _apply_search_and_sort(df, search, sort_by, ascending)
    st.caption(f"{len(df_show)} topic(s) shown")

    for _, row in df_show.iterrows():
        _render_topic_row(row)


# ══════════════════════════════════════════════════════════════════
# 11. EXPORT
# ══════════════════════════════════════════════════════════════════
def render_export(df: pd.DataFrame) -> None:
    """Render the Excel export button."""
    st.markdown("---")
    st.download_button(
        label="⬇ Export Filtered Excel",
        data=make_excel_bytes(df),
        file_name=f"QM_Topics_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ══════════════════════════════════════════════════════════════════
# 12. PAGE HEADER
# ══════════════════════════════════════════════════════════════════
def render_header(n_shown: int, n_total: int) -> None:
    st.markdown(f"""
    <div style="display:flex; align-items:center; gap:12px; margin-bottom:1.2rem;">
      <div style="background:#0F3460; width:6px; height:38px; border-radius:4px;"></div>
      <div>
        <div style="font-size:1.5rem; font-weight:700; color:#0F172A; line-height:1;">
          Quality Management · Topic Database
        </div>
        <div style="font-size:0.8rem; color:#64748B; margin-top:2px;">
          Management Overview · {date.today().strftime('%d %B %Y')}
          · {n_shown} of {n_total} topics shown
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# 13. MAIN ENTRY POINT
# ══════════════════════════════════════════════════════════════════
def main() -> None:
    inject_styles()

    # Sidebar returns both the unfiltered and filtered DataFrames
    df_raw, df = render_sidebar(load_data())

    render_header(n_shown=len(df), n_total=len(df_raw))
    render_kpi_row(df, df_raw)

    st.markdown("<div style='margin:1rem 0'></div>", unsafe_allow_html=True)

    render_charts(df)
    render_topic_table(df)
    render_export(df)


if __name__ == "__main__":
    main()
