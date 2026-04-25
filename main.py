import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import io

# ── Page config ──────────────────────────────────────────────────
st.set_page_config(
    page_title="QM Topic Dashboard",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Global styles ────────────────────────────────────────────────
st.markdown("""
<style>
/* Inter font */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Main background */
.main .block-container { padding: 1.5rem 2rem; background: #F8FAFC; }

/* Metric cards */
.kpi-card {
    background:gray;
    border-radius: 12px;
    padding: 1.1rem 1.2rem;
    border-left: 5px solid;
    box-shadow: 0 1px 6px rgba(0,0,0,0.07);
    margin-bottom: 0.4rem;
}
.kpi-label { font-size: 0.72rem; font-weight: 600; color: #64748B;
             text-transform: uppercase; letter-spacing: 0.06em; }
.kpi-value { font-size: 2.2rem; font-weight: 700; line-height: 1.1; margin-top: 2px; }
.kpi-sub   { font-size: 0.72rem; color: #94A3B8; margin-top: 2px; }

/* Section headers */
.section-title {
    font-size: 0.8rem; font-weight: 700; color: #334155;
    text-transform: uppercase; letter-spacing: 0.08em;
    border-bottom: 2px solid #E2E8F0; padding-bottom: 6px; margin: 0.6rem 0 0.8rem;
}

/* Severity badges */
.badge {
    display: inline-block; padding: 2px 10px; border-radius: 20px;
    font-size: 0.72rem; font-weight: 600; letter-spacing: 0.03em;
}
.badge-critical { background:#FEE2E2; color:#B91C1C; }
.badge-high     { background:#FEE2E2; color:#DC2626; }
.badge-medium   { background:#FEF3C7; color:#D97706; }
.badge-low      { background:#DCFCE7; color:#15803D; }
.badge-open     { background:#DBEAFE; color:#1D4ED8; }
.badge-progress { background:#FEF3C7; color:#D97706; }
.badge-blocked  { background:#FEE2E2; color:#DC2626; }
.badge-closed   { background:#DCFCE7; color:#15803D; }
.badge-yes      { background:#FEE2E2; color:#B91C1C; }
.badge-no       { background:#F1F5F9; color:#64748B; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: #0F3460;
}
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stMultiSelect label { color: #CBD5E1 !important; }

/* Detail expander */
.detail-row { display: flex; gap: 1rem; margin-bottom: 0.5rem; }
.detail-label { font-size: 0.75rem; color: #64748B; font-weight: 600;
                text-transform: uppercase; min-width: 140px; }
.detail-value { font-size: 0.85rem; color: #1E293B; }

/* Scrollable table */
.table-container { overflow-x: auto; }
</style>
""", unsafe_allow_html=True)


# ── Data loading ─────────────────────────────────────────────────
@st.cache_data
def load_data(uploaded=None):
    if uploaded:
        df = pd.read_excel(uploaded, sheet_name="Topic Database", header=1)
    else:
        # Embedded sample data (the 8 rows from your file)
        rows = [
            dict(ID=1, **{"Topic Group":"Welding Defects","Sub-Topic":"Porosity in MIG Welds",
                 "Category":"Elements","Severity":"High","Opening Date":"2026-02-24",
                 "Close Date":None,"PIC NED":"J. Müller","PIC HQ":"J. Müller","Status":"Open",
                 "Cust. Impact":"Yes","Days Open":38,"Aging Bucket":"0–3 Months",
                 "Escalated":"YES","Problem Description":"Porosity in batch WD-22; ~15% rejection rate.",
                 "Root Cause Analysis":"Shielding gas contamination on line 3.",
                 "Corrective Actions":"Replaced gas supply line; tightened fittings.",
                 "Prevention of recurrence":"Replaced gas supply line; tightened fittings.",
                 "Next Steps":"Requalification test next week.","Milestones / Dates":"8D Due: 2026-03-23"}),
            dict(ID=2, **{"Topic Group":"Assembly Sequence","Sub-Topic":"Bolt torque out-of-spec",
                 "Category":"Assembly","Severity":"Medium","Opening Date":"2025-11-06",
                 "Close Date":"2026-03-06","PIC NED":"A. Schmidt","PIC HQ":"A. Schmidt","Status":"Closed",
                 "Cust. Impact":"No","Days Open":120,"Aging Bucket":"3–6 Months",
                 "Escalated":"no","Problem Description":"Torque 10% below spec on rear bracket.",
                 "Root Cause Analysis":"Calibration drift on station 7.",
                 "Corrective Actions":"Wrench recalibrated; audit done.",
                 "Prevention of recurrence":"Wrench recalibrated; audit done.",
                 "Next Steps":None,"Milestones / Dates":None}),
            dict(ID=3, **{"Topic Group":"Supplier Quality","Sub-Topic":"Dimensional deviation – X401",
                 "Category":"Cross","Severity":"Critical","Opening Date":"2025-07-09",
                 "Close Date":None,"PIC NED":"L. Bauer","PIC HQ":"L. Bauer","Status":"Blocked",
                 "Cust. Impact":"Yes","Days Open":268,"Aging Bucket":"6–12 Months",
                 "Escalated":"YES","Problem Description":"OD of X401 exceeds tolerance +0.3 mm.",
                 "Root Cause Analysis":"Under investigation – supplier audit planned.",
                 "Corrective Actions":"Interim: 100% incoming inspection.",
                 "Prevention of recurrence":"Interim: 100% incoming inspection.",
                 "Next Steps":"Awaiting supplier 8D response.","Milestones / Dates":"Supplier response due: 2026-03-19"}),
            dict(ID=4, **{"Topic Group":"Paint & Coating","Sub-Topic":"Surface adhesion failure",
                 "Category":"Elements","Severity":"High","Opening Date":"2025-02-09",
                 "Close Date":None,"PIC NED":"K. Vogel","PIC HQ":"K. Vogel","Status":"Open",
                 "Cust. Impact":"No","Days Open":418,"Aging Bucket":"> 1 Year",
                 "Escalated":"YES","Problem Description":"Peeling after 48h salt-spray test.",
                 "Root Cause Analysis":"Pre-treatment bath concentration out of range.",
                 "Corrective Actions":"Bath replenished; batch quarantined.",
                 "Prevention of recurrence":"Bath replenished; batch quarantined.",
                 "Next Steps":"Retest batch after rework.","Milestones / Dates":None}),
            dict(ID=5, **{"Topic Group":"Welding Defects","Sub-Topic":"Undercut on fillet welds",
                 "Category":"Elements","Severity":"Medium","Opening Date":"2026-01-30",
                 "Close Date":None,"PIC NED":"J. Müller","PIC HQ":"J. Müller","Status":"In Progress",
                 "Cust. Impact":"No","Days Open":63,"Aging Bucket":"0–3 Months",
                 "Escalated":"no","Problem Description":"Undercut >0.5 mm on fillet joints zone B.",
                 "Root Cause Analysis":"Travel speed too high; welder technique.",
                 "Corrective Actions":"Additional welder training completed.",
                 "Prevention of recurrence":"Additional welder training completed.",
                 "Next Steps":"Monitor next 3 production runs.","Milestones / Dates":None}),
            dict(ID=6, **{"Topic Group":"Supplier Quality","Sub-Topic":"Late delivery – Component Y7",
                 "Category":"Cross","Severity":"Medium","Opening Date":"2026-01-15",
                 "Close Date":None,"PIC NED":"M. Weber","PIC HQ":"M. Weber","Status":"Open",
                 "Cust. Impact":"Yes","Days Open":78,"Aging Bucket":"0–3 Months",
                 "Escalated":"no","Problem Description":"Supplier 3-5 days late consistently.",
                 "Root Cause Analysis":"Raw material shortage at supplier.",
                 "Corrective Actions":"Dual-sourcing approval in progress.",
                 "Prevention of recurrence":"Dual-sourcing approval in progress.",
                 "Next Steps":"Qualify second supplier by month-end.","Milestones / Dates":None}),
            dict(ID=7, **{"Topic Group":"Assembly Sequence","Sub-Topic":"Misaligned bracket – Stn 4",
                 "Category":"Assembly","Severity":"Low","Opening Date":"2026-03-01",
                 "Close Date":None,"PIC NED":"A. Schmidt","PIC HQ":"A. Schmidt","Status":"In Progress",
                 "Cust. Impact":"No","Days Open":33,"Aging Bucket":"0–3 Months",
                 "Escalated":"no","Problem Description":"Bracket deviation 2mm on 8% of parts.",
                 "Root Cause Analysis":"Fixture wear on station 4.",
                 "Corrective Actions":"Fixture replaced and re-qualified.",
                 "Prevention of recurrence":"Fixture replaced and re-qualified.",
                 "Next Steps":"Monitor next 3 runs.","Milestones / Dates":None}),
            dict(ID=8, **{"Topic Group":"Paint & Coating","Sub-Topic":"Orange peel texture",
                 "Category":"Elements","Severity":"Low","Opening Date":"2025-08-28",
                 "Close Date":None,"PIC NED":"K. Vogel","PIC HQ":"K. Vogel","Status":"Blocked",
                 "Cust. Impact":"No","Days Open":218,"Aging Bucket":"6–12 Months",
                 "Escalated":"YES","Problem Description":"Orange peel on exterior panels.",
                 "Root Cause Analysis":"Paint viscosity out of spec.",
                 "Corrective Actions":"Viscosity adjusted; process frozen.",
                 "Prevention of recurrence":"Viscosity adjusted; process frozen.",
                 "Next Steps":"Waiting for customer waiver.","Milestones / Dates":None}),
        ]
        df = pd.DataFrame(rows)

    # Normalise
    df.columns = [c.strip() for c in df.columns]
    df = df.dropna(subset=["ID"])
    df["ID"] = df["ID"].astype(int)
    df["Escalated"] = df["Escalated"].astype(str).str.strip().str.upper().map(
        lambda x: "YES" if x == "YES" else "No")
    df["Cust. Impact"] = df.get("Cust. Impact", pd.Series(["No"]*len(df))).fillna("No")
    for d_col in ["Opening Date", "Close Date"]:
        if d_col in df.columns:
            df[d_col] = pd.to_datetime(df[d_col], errors="coerce")
    df["Days Open"] = pd.to_numeric(df.get("Days Open", 0), errors="coerce").fillna(0).astype(int)
    return df.sort_values("ID").reset_index(drop=True)


# ── Colour helpers ────────────────────────────────────────────────
SEV_ORDER  = ["Critical","High","Medium","Low"]
STAT_ORDER = ["Open","In Progress","Blocked","Closed"]
SEV_COLOR  = {"Critical":"#DC2626","High":"#EF4444","Medium":"#F59E0B","Low":"#22C55E"}
STAT_COLOR = {"Open":"#3B82F6","In Progress":"#F59E0B","Blocked":"#EF4444","Closed":"#22C55E"}
CAT_COLOR  = {"Elements":"#3B82F6","Assembly":"#22C55E","Cross":"#F59E0B"}

def sev_badge(v):
    m = {"Critical":"critical","High":"high","Medium":"medium","Low":"low"}
    cls = m.get(v, "no")
    return f'<span class="badge badge-{cls}">{v}</span>'

def stat_badge(v):
    m = {"Open":"open","In Progress":"progress","Blocked":"blocked","Closed":"closed"}
    cls = m.get(v, "no")
    return f'<span class="badge badge-{cls}">{v}</span>'

def esc_badge(v):
    cls = "yes" if v == "YES" else "no"
    return f'<span class="badge badge-{cls}">{v}</span>'


# ══════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("##  QM Dashboard")
    st.markdown("---")

    uploaded = st.file_uploader("Upload Excel", type=["xlsx"],
                                help="Upload your Topic Database Excel file")
    st.markdown("---")
    st.markdown("### Filters")

    df_raw = load_data(uploaded)

    sel_cat = st.multiselect("Category",
        options=sorted(df_raw["Category"].dropna().unique()),
        default=sorted(df_raw["Category"].dropna().unique()))

    sel_stat = st.multiselect("Status",
        options=STAT_ORDER,
        default=STAT_ORDER)

    sel_esc = st.multiselect("Escalated",
        options=["YES","No"], default=["YES","No"])

    pic_opts = sorted(df_raw["PIC NED"].dropna().unique())
    sel_pic  = st.multiselect("PIC", options=pic_opts, default=pic_opts)

    days_max = int(df_raw["Days Open"].max()) if len(df_raw) else 500
    sel_days = st.slider("Max. Days Open", 0, max(days_max, 500), max(days_max, 500))

    st.markdown("---")
    st.caption(f" {date.today().strftime('%d %b %Y')}")


# ── Apply filters ─────────────────────────────────────────────────
df = df_raw.copy()
if sel_cat:   df = df[df["Category"].isin(sel_cat)]
if sel_stat:  df = df[df["Status"].isin(sel_stat)]
if sel_esc:   df = df[df["Escalated"].isin(sel_esc)]
if sel_pic:   df = df[df["PIC NED"].isin(sel_pic)]
df = df[df["Days Open"] <= sel_days]


# ══════════════════════════════════════════════════════════════════
# MAIN CONTENT
# ══════════════════════════════════════════════════════════════════
st.markdown(f"""
<div style="display:flex; align-items:center; gap:12px; margin-bottom:1.2rem;">
  <div style="background:#0F3460; width:6px; height:38px; border-radius:4px;"></div>
  <div>
    <div style="font-size:1.5rem; font-weight:700; color:#0F172A; line-height:1;">
      Quality Management · Topic Database
    </div>
    <div style="font-size:0.8rem; color:#64748B; margin-top:2px;">
      Management Overview · {date.today().strftime('%d %B %Y')} · {len(df)} of {len(df_raw)} topics shown
    </div>
  </div>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# KPI ROW
# ══════════════════════════════════════════════════════════════════
k1, k2, k3, k4, k5, k6 = st.columns(6)

def kpi(col, label, value, color, sub=""):
    col.markdown(f"""
    <div class="kpi-card" style="border-color:{color}">
      <div class="kpi-label">{label}</div>
      <div class="kpi-value" style="color:{color}">{value}</div>
      <div class="kpi-sub">{sub}</div>
    </div>""", unsafe_allow_html=True)

total   = len(df)
n_open  = len(df[df["Status"]=="Open"])
n_prog  = len(df[df["Status"]=="In Progress"])
n_blk   = len(df[df["Status"]=="Blocked"])
n_cls   = len(df[df["Status"]=="Closed"])
n_esc   = len(df[df["Escalated"]=="YES"])

kpi(k1, "Total Topics",   total,  "#0F3460", f"{len(df_raw)} total")
kpi(k2, "Open",           n_open, "#3B82F6", "active")
kpi(k3, "In Progress",    n_prog, "#F59E0B", "active")
kpi(k4, "Blocked",        n_blk,  "#EF4444", "needs action")
kpi(k5, "Closed",         n_cls,  "#22C55E", "resolved")
kpi(k6, "Escalated",      n_esc,  "#B91C1C", "⚠ mgmt attn")


st.markdown("<div style='margin:1rem 0'></div>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════
# CHARTS ROW
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">📊 Analytics</div>', unsafe_allow_html=True)
c1, c2, c3 = st.columns([1.3, 1, 1])

with c1:
    # Status × Category stacked bar
    grp = df.groupby(["Category","Status"]).size().reset_index(name="n")
    fig = px.bar(grp, x="Category", y="n", color="Status",
                 color_discrete_map=STAT_COLOR,
                 category_orders={"Status": STAT_ORDER},
                 text_auto=True,
                 labels={"n":"Topics","Category":""},
                 title="Topics by Category & Status")
    fig.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                      legend_title="", font_family="Inter",
                      margin=dict(t=40, b=10, l=10, r=10), height=280,
                      legend=dict(orientation="h", yanchor="bottom", y=1.02))
    fig.update_traces(textfont_size=11)
    st.plotly_chart(fig, use_container_width=True)

with c2:
    # Aging bucket bar
    aging_order = ["0–3 Months","3–6 Months","6–12 Months","> 1 Year"]
    aging = df["Aging Bucket"].value_counts().reindex(aging_order).dropna()
    colors_aging = ["#22C55E","#F59E0B","#EF4444","#B91C1C"]
    fig3 = go.Figure(go.Bar(
        x=aging.index, y=aging.values,
        marker_color=colors_aging[:len(aging)],
        text=aging.values, textposition="outside"))
    fig3.update_layout(title="Aging Distribution", plot_bgcolor="white",
                       paper_bgcolor="white", font_family="Inter",
                       showlegend=False, height=280,
                       margin=dict(t=40, b=10, l=10, r=10),
                       yaxis=dict(showgrid=True, gridcolor="#F1F5F9"),
                       xaxis=dict(title=""))
    st.plotly_chart(fig3, use_container_width=True)

#with c3:
    # PIC workload — open + in progress + blocked per person
    #active_df = df[df["Status"].isin(["Open","In Progress","Blocked"])]
    #pic_grp = active_df.groupby(["PIC NED","Status"]).size().reset_index(name="n")
    #fig4 = px.bar(pic_grp, x="PIC NED", y="n", color="Status",
                  #color_discrete_map=STAT_COLOR,
                  #category_orders={"Status": ["Open","In Progress","Blocked"]},
                  #text_auto=True,
                  #labels={"n":"Open Topics","PIC NED":""},
                 # title="Active Topics by PIC")
    #fig4.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                       #legend_title="", font_family="Inter",
                      # margin=dict(t=40, b=10, l=10, r=10), height=280,
                       #legend=dict(orientation="h", yanchor="bottom", y=1.02))
   # fig4.update_traces(textfont_size=11)
   # st.plotly_chart(fig4, use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# TOPIC TABLE
# ══════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">Topic Overview</div>', unsafe_allow_html=True)

# Search bar
search = st.text_input("Search topics, PICs, descriptions…", placeholder="Type to filter…", label_visibility="collapsed")
if search:
    mask = df.apply(lambda r: r.astype(str).str.contains(search, case=False, na=False).any(), axis=1)
    df_show = df[mask]
else:
    df_show = df

# Sort control
sort_col, sort_dir = st.columns([3, 1])
with sort_col:
    sort_by = st.selectbox("Sort by", ["ID","Days Open","Status","Category"],
                           label_visibility="collapsed")
with sort_dir:
    ascending = st.selectbox("Order", ["↑ Ascending","↓ Descending"],
                              label_visibility="collapsed") == "↑ Ascending"

sev_rank  = {"Critical":0,"High":1,"Medium":2,"Low":3}
stat_rank = {"Blocked":0,"Open":1,"In Progress":2,"Closed":3}

if sort_by == "Status":
    df_show = df_show.copy()
    df_show["_rank"] = df_show["Status"].map(stat_rank)
    df_show = df_show.sort_values("_rank", ascending=ascending).drop(columns="_rank")
else:
    df_show = df_show.sort_values(sort_by, ascending=ascending)

st.caption(f"{len(df_show)} topic(s) shown")

# Render table rows
for _, row in df_show.iterrows():
    esc_flag = row["Escalated"] == "YES"
    bg = "#FFF8F8" if esc_flag else "white"
    border = "2px solid #FCA5A5" if esc_flag else "1px solid #E2E8F0"
    days_color = "#DC2626" if row["Days Open"] > 180 else ("#F59E0B" if row["Days Open"] > 60 else "#64748B")

    with st.expander(
        f"#{row['ID']}  ·  {row['Sub-Topic']}  —  {row['Topic Group']}",
        expanded=False
    ):
        col_a, col_b, col_c, col_d = st.columns([1.2, 1.2, 1, 1])
        with col_a:
            st.markdown(f"**Category:** {row['Category']}")
            st.markdown(f"**Status:** {stat_badge(row['Status'])}", unsafe_allow_html=True)
            st.markdown(f"**Escalated:** {esc_badge(row['Escalated'])}", unsafe_allow_html=True)
        with col_b:
            st.markdown(f"**PIC NED:** {row['PIC NED']}")
            st.markdown(f"**PIC HQ:** {row.get('PIC HQ','—')}")
        with col_c:
            st.markdown(f"**Days Open:** :{('red' if row['Days Open']>180 else 'orange' if row['Days Open']>60 else 'gray')}[{row['Days Open']}d]")
            st.markdown(f"**Aging:** {row.get('Aging Bucket','—')}")
        with col_d:
            op = row['Opening Date'].strftime('%d %b %Y') if pd.notna(row.get('Opening Date')) else "—"
            cl = row['Close Date'].strftime('%d %b %Y') if pd.notna(row.get('Close Date')) else "Open"
            st.markdown(f"**Opened:** {op}")
            st.markdown(f"**Closed:** {cl}")
            if pd.notna(row.get('Milestones / Dates')) and str(row.get('Milestones / Dates', '')) not in ('', 'nan'):
                st.markdown(f"**Milestone:** {row['Milestones / Dates']}")

        st.markdown("---")
        c1d, c2d = st.columns(2)
        with c1d:
            st.markdown("**Problem Description**")
            st.info(row.get("Problem Description","—") or "—")
            st.markdown("**Root Cause Analysis**")
            st.info(row.get("Root Cause Analysis","—") or "—")
        with c2d:
            st.markdown("**Corrective Actions**")
            st.success(row.get("Corrective Actions","—") or "—")
            st.markdown("**Next Steps**")
            st.warning(row.get("Next Steps","—") or "—")

        prev = row.get("Prevention of recurrence","")
        if prev and str(prev) not in ("nan",""):
            st.markdown("**Prevention of Recurrence**")
            st.markdown(f"> {prev}")


# ══════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════
st.markdown("---")
ec1, ec2 = st.columns([1, 5])
with ec1:
    buf = io.BytesIO()
    df_show.to_excel(buf, index=False, sheet_name="Filtered Topics")
    st.download_button(
        "⬇ Export Filtered Excel",
        data=buf.getvalue(),
        file_name=f"QM_Topics_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
