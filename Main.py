
import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import json
import os

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Topic Database",
    page_icon="🗂️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ──────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }
header[data-testid="stHeader"] { background: transparent; }

.kpi-card {
    background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px;
    padding: 20px 24px; text-align: center;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06); transition: box-shadow .2s;
}
.kpi-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.10); }
.kpi-label { font-size: 12px; font-weight: 600; color: #6B7280; text-transform: uppercase; letter-spacing: .06em; margin-bottom: 6px; }
.kpi-value { font-size: 34px; font-weight: 700; color: #111827; line-height: 1; }
.kpi-sub   { font-size: 12px; color: #9CA3AF; margin-top: 4px; }

.section-title {
    font-size: 22px; font-weight: 700; color: #111827;
    margin: 0 0 16px 0; padding-bottom: 8px; border-bottom: 2px solid #E5E7EB;
}
.progress-outer { background:#E5E7EB; border-radius:99px; height:10px; overflow:hidden; margin-top:8px; }
.progress-inner { height:100%; border-radius:99px; background: linear-gradient(90deg,#3B82F6,#6366F1); transition: width .6s ease; }

.bucket-row  { display:flex; align-items:center; gap:10px; margin-bottom:8px; font-size:13px; }
.bucket-label{ width:140px; color:#374151; flex-shrink:0; }
.bucket-bar-outer{ flex:1; background:#F3F4F6; border-radius:99px; height:8px; overflow:hidden; }
.bucket-bar-inner{ height:100%; border-radius:99px; }
.bucket-count{ width:28px; text-align:right; color:#6B7280; font-size:12px; font-family:'DM Mono',monospace; }
</style>
""", unsafe_allow_html=True)

# ── Data persistence (JSON file) ────────────────────────────────────────────────
DATA_FILE = "topics_data.json"
CATEGORIES = ["Elements", "Assembly", "Cross"]
STATUSES   = ["Open", "In Progress", "Blocked", "Closed"]

def aging_bucket(opening_date):
    if pd.isna(opening_date):
        return "Unknown"
    days = (datetime.today() - pd.to_datetime(opening_date)).days
    if days <= 90:    return "0–3 Months"
    elif days <= 180: return "3–6 Months"
    elif days <= 365: return "6–12 Months"
    else:             return "> 1 Year"

def load_data() -> pd.DataFrame:
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r") as f:
                records = json.load(f)
            df = pd.DataFrame(records)
            df["opening_date"] = pd.to_datetime(df["opening_date"], errors="coerce")
            return df
        except Exception:
            pass
    # Seed with demo data
    today = datetime.today()
    return pd.DataFrame([
        {"id": 1, "topic_group": "Welding Defects", "sub_topic": "Porosity in MIG Welds",
         "category": "Elements", "opening_date": today - timedelta(days=20), "pic": "J. Müller",
         "problem_description": "Porosity observed in batch WD-22. Approx. 15% rejection rate.",
         "root_cause_analysis": "Shielding gas contamination detected in line 3.",
         "corrective_actions": "Replaced gas supply line; tightened fittings.",
         "status": "In Progress", "next_steps": "Requalification test scheduled for next week.",
         "customer_impact": True, "pictures": []},
        {"id": 2, "topic_group": "Assembly Sequence", "sub_topic": "Bolt torque out-of-spec",
         "category": "Assembly", "opening_date": today - timedelta(days=130), "pic": "A. Schmidt",
         "problem_description": "Torque values 10% below spec on rear bracket.",
         "root_cause_analysis": "Calibration drift on station 7 torque wrench.",
         "corrective_actions": "Wrench recalibrated; process audit completed.",
         "status": "Closed", "next_steps": "", "customer_impact": False, "pictures": []},
        {"id": 3, "topic_group": "Supplier Quality", "sub_topic": "Dimensional deviation – Part X401",
         "category": "Cross", "opening_date": today - timedelta(days=250), "pic": "L. Bauer",
         "problem_description": "OD of X401 exceeds tolerance +0.3 mm consistently.",
         "root_cause_analysis": "Under investigation – supplier audit planned.",
         "corrective_actions": "Interim: 100% inspection at receiving.",
         "status": "Blocked", "next_steps": "Awaiting supplier response to 8D report.",
         "customer_impact": True, "pictures": []},
        {"id": 4, "topic_group": "Paint & Coating", "sub_topic": "Surface adhesion failure",
         "category": "Elements", "opening_date": today - timedelta(days=400), "pic": "K. Vogel",
         "problem_description": "Peeling detected after 48h salt-spray test.",
         "root_cause_analysis": "Pre-treatment bath concentration out of range.",
         "corrective_actions": "Bath replenished; batch quarantined.",
         "status": "Open", "next_steps": "Retest batch after rework.",
         "customer_impact": False, "pictures": []},
    ])

def save_data(df: pd.DataFrame):
    records = df.copy()
    records["opening_date"] = records["opening_date"].astype(str)
    with open(DATA_FILE, "w") as f:
        json.dump(records.to_dict(orient="records"), f, indent=2, default=str)

def next_id(df):
    return int(df["id"].max() + 1) if len(df) > 0 else 1

# ── Session state ───────────────────────────────────────────────────────────────
if "df" not in st.session_state:
    st.session_state.df = load_data()
if "edit_id" not in st.session_state:
    st.session_state.edit_id = None

df: pd.DataFrame = st.session_state.df
df["_bucket"] = df["opening_date"].apply(aging_bucket)

cat_colors    = {"Elements": "#3B82F6", "Assembly": "#10B981", "Cross": "#F59E0B"}
bucket_colors = {"0–3 Months": "#6EE7B7", "3–6 Months": "#FCD34D", "6–12 Months": "#FCA5A5", "> 1 Year": "#F87171"}
status_colors = {"Open": "#EF4444", "In Progress": "#3B82F6", "Blocked": "#F59E0B", "Closed": "#10B981"}

# ── Sidebar ─────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🗂️ Topic Database")
    st.markdown("---")
    nav = st.radio("Navigation", ["📊 Dashboard", "📋 All Topics", "➕ New Topic", "📤 Export"],
                   label_visibility="collapsed")
    st.markdown("---")
    st.markdown("**Filter Topics**")
    f_cat    = st.multiselect("Category",  CATEGORIES, default=CATEGORIES)
    f_status = st.multiselect("Status",    STATUSES,   default=STATUSES)
    f_search = st.text_input("🔍 Search")
    st.markdown("---")
    st.caption(f"Total topics in DB: **{len(df)}**")

# Apply filters
mask = df["category"].isin(f_cat) & df["status"].isin(f_status)
if f_search:
    q = f_search.lower()
    mask &= (df["topic_group"].str.lower().str.contains(q, na=False) |
             df["problem_description"].str.lower().str.contains(q, na=False) |
             df["sub_topic"].str.lower().str.contains(q, na=False))
dff = df[mask].copy()

# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if "Dashboard" in nav:
    st.markdown('<p class="section-title">📊 Summary Dashboard</p>', unsafe_allow_html=True)

    total     = len(df)
    open_n    = len(df[df["status"] == "Open"])
    closed_n  = len(df[df["status"] == "Closed"])
    prog_n    = len(df[df["status"] == "In Progress"])
    blocked_n = len(df[df["status"] == "Blocked"])
    pct       = round(closed_n / total * 100) if total else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    for col, label, val in [
        (c1,"Total Topics", total),   (c2,"Open", open_n),
        (c3,"In Progress",  prog_n),  (c4,"Blocked", blocked_n),
        (c5,"Closed",       closed_n),(c6,"Completion %", f"{pct}%"),
    ]:
        col.markdown(f"""
        <div class="kpi-card">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{val}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown(f"""
    <div style="font-size:13px;color:#374151;font-weight:600;">Overall Completion</div>
    <div class="progress-outer"><div class="progress-inner" style="width:{pct}%"></div></div>
    <div style="font-size:11px;color:#9CA3AF;margin-top:4px;">{closed_n} of {total} topics closed</div>
    """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**⏱ Time Aging Buckets (non-closed topics)**")
        bucket_order = ["0–3 Months", "3–6 Months", "6–12 Months", "> 1 Year"]
        active = df[df["status"] != "Closed"].copy()
        active["_bucket"] = active["opening_date"].apply(aging_bucket)
        bucket_counts = active["_bucket"].value_counts()
        max_b = max(bucket_counts.values) if len(bucket_counts) else 1
        for b in bucket_order:
            cnt = bucket_counts.get(b, 0)
            pct_b = int(cnt / max_b * 100) if max_b else 0
            color = bucket_colors.get(b, "#D1D5DB")
            st.markdown(f"""
            <div class="bucket-row">
              <div class="bucket-label">{b}</div>
              <div class="bucket-bar-outer"><div class="bucket-bar-inner" style="width:{pct_b}%;background:{color}"></div></div>
              <div class="bucket-count">{cnt}</div>
            </div>""", unsafe_allow_html=True)

    with col_right:
        st.markdown("**🏷 Category Breakdown**")
        for cat in CATEGORIES:
            sub_df  = df[df["category"] == cat]
            open_c  = len(sub_df[sub_df["status"]=="Open"])
            prog_c  = len(sub_df[sub_df["status"]=="In Progress"])
            blk_c   = len(sub_df[sub_df["status"]=="Blocked"])
            clos_c  = len(sub_df[sub_df["status"]=="Closed"])
            total_c = len(sub_df)
            color   = cat_colors.get(cat, "#6B7280")
            pct_c   = round(clos_c / total_c * 100) if total_c else 0
            st.markdown(f"""
            <div style="border:1px solid #E5E7EB;border-left:4px solid {color};
                 border-radius:8px;padding:12px 16px;margin-bottom:10px;background:#FAFAFA;">
              <div style="font-weight:700;color:#111827;font-size:14px;">{cat}</div>
              <div style="font-size:12px;color:#6B7280;margin-top:4px;">
                Total: <b>{total_c}</b> &nbsp;|&nbsp;
                Open: <b style="color:#EF4444">{open_c}</b> &nbsp;|&nbsp;
                In Progress: <b style="color:#3B82F6">{prog_c}</b> &nbsp;|&nbsp;
                Blocked: <b style="color:#F59E0B">{blk_c}</b> &nbsp;|&nbsp;
                Closed: <b style="color:#10B981">{clos_c}</b>
              </div>
              <div class="progress-outer" style="margin-top:8px;">
                <div class="progress-inner" style="width:{pct_c}%;background:{color}"></div>
              </div>
            </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# ALL TOPICS
# ══════════════════════════════════════════════════════════════════════════════
elif "All Topics" in nav:
    st.markdown('<p class="section-title">📋 All Topics</p>', unsafe_allow_html=True)
    st.caption(f"Showing **{len(dff)}** of **{len(df)}** topics")

    if dff.empty:
        st.info("No topics match the current filters.")
    else:
        for _, row in dff.sort_values("opening_date", ascending=False).iterrows():
            with st.expander(
                f"[#{int(row['id'])}] {row['topic_group']} — {row['sub_topic']}  |  {row['category']}  |  {row['status']}",
                expanded=False,
            ):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Category",     row["category"])
                c2.metric("Status",       row["status"])
                c3.metric("PIC",          row.get("pic","—"))
                od = row["opening_date"]
                c4.metric("Opening Date", str(od.date()) if not pd.isna(od) else "—")

                st.markdown(f"**📝 Problem Description**  \n{row.get('problem_description','—')}")
                st.markdown(f"**🔍 Root Cause Analysis**  \n{row.get('root_cause_analysis','—')}")
                st.markdown(f"**🔧 Corrective Actions**  \n{row.get('corrective_actions','—')}")
                st.markdown(f"**🔜 Next Steps**  \n{row.get('next_steps','—')}")
                st.markdown(f"**Customer Impact**: {'✅ Yes' if row.get('customer_impact') else '❌ No'}")
                st.markdown(f"**Aging Bucket**: {row['_bucket']}")

                if st.button(f"✏️ Edit Topic #{int(row['id'])}", key=f"edit_{row['id']}"):
                    st.session_state.edit_id = int(row["id"])

    # Inline edit form
    if st.session_state.edit_id is not None:
        eid = st.session_state.edit_id
        row_match = df[df["id"] == eid]
        if len(row_match) > 0:
            row = row_match.iloc[0]
            st.markdown("---")
            st.markdown(f"### ✏️ Edit Topic #{eid}")
            with st.form(f"edit_form_{eid}"):
                col1, col2 = st.columns(2)
                tg   = col1.text_input("Topic Group",   value=row["topic_group"])
                st_  = col1.text_input("Sub-Topic",      value=row["sub_topic"])
                cat  = col1.selectbox("Category",  CATEGORIES, index=CATEGORIES.index(row["category"]))
                pic  = col2.text_input("PIC",            value=row["pic"])
                stat = col2.selectbox("Status",    STATUSES,   index=STATUSES.index(row["status"]))
                od   = col2.date_input("Opening Date",
                                       value=row["opening_date"].date() if not pd.isna(row["opening_date"]) else date.today())
                ci   = col2.checkbox("Customer Impact", value=bool(row.get("customer_impact", False)))
                pd_  = st.text_area("Problem Description",  value=row.get("problem_description",""), height=80)
                rca  = st.text_area("Root Cause Analysis",  value=row.get("root_cause_analysis",""),  height=80)
                ca   = st.text_area("Corrective Actions",   value=row.get("corrective_actions",""),   height=80)
                ns   = st.text_area("Next Steps",            value=row.get("next_steps",""),           height=60)
                col_a, col_b = st.columns(2)
                save  = col_a.form_submit_button("💾 Save Changes", type="primary")
                cancel= col_b.form_submit_button("❌ Cancel")
                if save:
                    idx = df.index[df["id"] == eid][0]
                    for field, val in [
                        ("topic_group", tg), ("sub_topic", st_), ("category", cat),
                        ("pic", pic), ("status", stat), ("opening_date", pd.Timestamp(od)),
                        ("customer_impact", ci), ("problem_description", pd_),
                        ("root_cause_analysis", rca), ("corrective_actions", ca), ("next_steps", ns),
                    ]:
                        df.at[idx, field] = val
                    st.session_state.df = df
                    save_data(df)
                    st.session_state.edit_id = None
                    st.success("Topic updated!")
                    st.rerun()
                if cancel:
                    st.session_state.edit_id = None
                    st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# NEW TOPIC
# ══════════════════════════════════════════════════════════════════════════════
elif "New Topic" in nav:
    st.markdown('<p class="section-title">➕ New Topic</p>', unsafe_allow_html=True)
    with st.form("new_topic_form"):
        col1, col2 = st.columns(2)
        tg   = col1.text_input("Topic Group *")
        st_  = col1.text_input("Sub-Topic *")
        cat  = col1.selectbox("Category *", CATEGORIES)
        pic  = col2.text_input("Person in Charge (PIC)")
        stat = col2.selectbox("Status *", STATUSES)
        od   = col2.date_input("Opening Date", value=date.today())
        ci   = col2.checkbox("Customer Impact")
        pd_  = st.text_area("Problem Description", height=80)
        rca  = st.text_area("Root Cause Analysis",  height=80)
        ca   = st.text_area("Corrective Actions",   height=80)
        ns   = st.text_area("Next Steps",            height=60)
        uploaded_pics = st.file_uploader("📎 Attach Pictures (optional)",
                                         accept_multiple_files=True, type=["png","jpg","jpeg"])
        submitted = st.form_submit_button("✅ Save Topic", type="primary")
        if submitted:
            if not tg or not st_:
                st.error("Please fill in Topic Group and Sub-Topic.")
            else:
                new_row = {
                    "id": next_id(df), "topic_group": tg, "sub_topic": st_,
                    "category": cat, "opening_date": pd.Timestamp(od), "pic": pic,
                    "problem_description": pd_, "root_cause_analysis": rca,
                    "corrective_actions": ca, "status": stat, "next_steps": ns,
                    "customer_impact": ci,
                    "pictures": [f.name for f in uploaded_pics] if uploaded_pics else [],
                }
                new_df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                st.session_state.df = new_df
                save_data(new_df)
                st.success(f"✅ Topic #{new_row['id']} saved successfully!")
                st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════════════
elif "Export" in nav:
    st.markdown('<p class="section-title">📤 Export Report</p>', unsafe_allow_html=True)

    export_all = st.toggle("Export ALL topics (ignore filters)", value=True)
    export_df  = df.copy() if export_all else dff.copy()
    st.info(f"Will export **{len(export_df)}** topics.")

    # ─ Excel ─────────────────────────────────────────────────────────────────
    st.subheader("📊 Excel Export")

    def build_excel(data: pd.DataFrame) -> bytes:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()

        # Styles
        HDR    = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        HDR2   = Font(name="Calibri", size=11, bold=True, color="1F3864")
        NORM   = Font(name="Calibri", size=10)
        BOLDF  = Font(name="Calibri", size=10, bold=True)
        BLUE_F = PatternFill("solid", fgColor="1F3864")
        MED_F  = PatternFill("solid", fgColor="2E75B6")
        LBL_F  = PatternFill("solid", fgColor="D9E1F2")
        CAT_F  = {"Elements": PatternFill("solid", fgColor="DBEAFE"),
                  "Assembly": PatternFill("solid", fgColor="D1FAE5"),
                  "Cross":    PatternFill("solid", fgColor="FEF3C7")}
        STA_F  = {"Open":        PatternFill("solid", fgColor="FEE2E2"),
                  "In Progress": PatternFill("solid", fgColor="DBEAFE"),
                  "Blocked":     PatternFill("solid", fgColor="FEF3C7"),
                  "Closed":      PatternFill("solid", fgColor="D1FAE5")}
        thin   = Side(style="thin", color="B0B0B0")
        brd    = Border(left=thin, right=thin, top=thin, bottom=thin)

        def hdr_row(ws, r, text, cols=8):
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
            c = ws.cell(row=r, column=1, value=text)
            c.font = HDR; c.fill = BLUE_F
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[r].height = 22

        def wcell(ws, r, col, val, bold=False, fill=None, align="left"):
            c = ws.cell(row=r, column=col, value=val)
            c.font = BOLDF if bold else NORM
            if fill: c.fill = fill
            c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
            c.border = brd
            return c

        # ── Sheet 1: Summary ──────────────────────────────────────────────────
        ws = wb.active; ws.title = "Summary Dashboard"
        hdr_row(ws, 1, "TOPIC DATABASE – SUMMARY DASHBOARD", cols=8)

        total_n  = len(data)
        closed_n = len(data[data["status"]=="Closed"])
        kpi_lbl  = ["Total Topics","Open","In Progress","Blocked","Closed","Completion %"]
        kpi_val  = [total_n,
                    len(data[data["status"]=="Open"]),
                    len(data[data["status"]=="In Progress"]),
                    len(data[data["status"]=="Blocked"]),
                    closed_n,
                    f"{round(closed_n/total_n*100) if total_n else 0}%"]
        ws.cell(row=3, column=1, value="KPI Overview").font = HDR2
        for i, (l, v) in enumerate(zip(kpi_lbl, kpi_val)):
            wcell(ws, 4, i+1, l, bold=True, fill=LBL_F, align="center")
            wcell(ws, 5, i+1, v, align="center")

        ws.cell(row=7, column=1, value="Aging Buckets (non-closed)").font = HDR2
        b_order  = ["0–3 Months","3–6 Months","6–12 Months","> 1 Year"]
        active   = data[data["status"] != "Closed"].copy()
        active["_bucket"] = active["opening_date"].apply(aging_bucket)
        b_counts = active["_bucket"].value_counts()
        for i, b in enumerate(b_order):
            wcell(ws, 8, i+1, b,                      bold=True, fill=LBL_F, align="center")
            wcell(ws, 9, i+1, int(b_counts.get(b,0)), align="center")

        ws.cell(row=11, column=1, value="Category Breakdown").font = HDR2
        for i, h in enumerate(["Category","Total","Open","In Progress","Blocked","Closed","% Done"]):
            wcell(ws, 12, i+1, h, bold=True, fill=LBL_F, align="center")
        for ri, cat in enumerate(CATEGORIES):
            sub = data[data["category"]==cat]; t = len(sub); cl = len(sub[sub["status"]=="Closed"])
            row_n = 13+ri
            for ci_, v in enumerate([cat, t, len(sub[sub["status"]=="Open"]),
                                      len(sub[sub["status"]=="In Progress"]),
                                      len(sub[sub["status"]=="Blocked"]), cl,
                                      f"{round(cl/t*100) if t else 0}%"]):
                wcell(ws, row_n, ci_+1, v, align="center").fill = CAT_F.get(cat, PatternFill())

        for i, w in enumerate([20,12,14,14,12,12,12,12], 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        # ── Sheet 2: Topics ───────────────────────────────────────────────────
        ws2 = wb.create_sheet("Topic Documentation")
        cols = [("ID",6),("Topic Group",22),("Sub-Topic",25),("Category",14),
                ("Opening Date",14),("PIC",16),("Status",14),
                ("Problem Description",40),("Root Cause Analysis",40),
                ("Corrective Actions",40),("Next Steps",30),
                ("Customer Impact",16),("Aging Bucket",16)]
        hdr_row(ws2, 1, "TOPIC DOCUMENTATION", cols=len(cols))
        for ci_, (name, width) in enumerate(cols, 1):
            c = ws2.cell(row=2, column=ci_, value=name)
            c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            c.fill = MED_F
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = brd
            ws2.column_dimensions[get_column_letter(ci_)].width = width

        data_s = data.sort_values("id").copy()
        data_s["_bucket"] = data_s["opening_date"].apply(aging_bucket)
        for ri, (_, r_) in enumerate(data_s.iterrows(), start=3):
            od = r_["opening_date"]
            vals = [int(r_["id"]), r_["topic_group"], r_["sub_topic"], r_["category"],
                    od.date() if not pd.isna(od) else "",
                    r_.get("pic",""), r_["status"],
                    r_.get("problem_description",""), r_.get("root_cause_analysis",""),
                    r_.get("corrective_actions",""), r_.get("next_steps",""),
                    "Yes" if r_.get("customer_impact") else "No", r_["_bucket"]]
            for ci_, val in enumerate(vals, 1):
                c = ws2.cell(row=ri, column=ci_, value=val)
                c.font = NORM
                c.alignment = Alignment(vertical="top", wrap_text=True)
                c.border = brd
            # Category color on col 4, Status color on col 7
            ws2.cell(row=ri, column=4).fill = CAT_F.get(r_["category"], PatternFill())
            ws2.cell(row=ri, column=7).fill = STA_F.get(r_["status"],   PatternFill())
            ws2.row_dimensions[ri].height = 40
        ws2.freeze_panes = "A3"

        buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

    xl = build_excel(export_df)
    st.download_button("⬇️ Download Excel Report", data=xl,
                       file_name=f"topic_report_{datetime.today().strftime('%Y%m%d')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary")

    # ─ PDF ────────────────────────────────────────────────────────────────────
    st.subheader("📄 PDF Export")

    def build_pdf(data: pd.DataFrame) -> bytes | None:
        try:
            from reportlab.lib.pagesizes import A4, landscape
            from reportlab.lib import colors
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import cm
            from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                            Table, TableStyle, PageBreak, HRFlowable)
            from reportlab.lib.enums import TA_CENTER

            buf = io.BytesIO()
            doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                                    leftMargin=1.5*cm, rightMargin=1.5*cm,
                                    topMargin=2*cm, bottomMargin=2*cm)

            DARK_BLUE = colors.HexColor("#1F3864")
            MED_BLUE  = colors.HexColor("#2E75B6")
            LT_BLUE   = colors.HexColor("#D9E1F2")
            GREEN     = colors.HexColor("#D1FAE5")
            YELLOW    = colors.HexColor("#FEF3C7")
            CAT_C  = {"Elements": LT_BLUE, "Assembly": GREEN, "Cross": YELLOW}
            STAT_C = {"Open": colors.HexColor("#FEE2E2"), "In Progress": colors.HexColor("#DBEAFE"),
                      "Blocked": YELLOW, "Closed": GREEN}

            title_s = ParagraphStyle("t", fontSize=18, textColor=DARK_BLUE,
                                     fontName="Helvetica-Bold", spaceAfter=6)
            sub_s   = ParagraphStyle("s", fontSize=10, textColor=colors.HexColor("#6B7280"),
                                     fontName="Helvetica", spaceAfter=14)
            h2_s    = ParagraphStyle("h2", fontSize=13, textColor=DARK_BLUE,
                                     fontName="Helvetica-Bold", spaceAfter=6, spaceBefore=14)
            cell_s  = ParagraphStyle("c", fontSize=7.5, fontName="Helvetica",
                                     leading=10, wordWrap="LTR")
            hdr_s   = ParagraphStyle("hd", fontSize=8, fontName="Helvetica-Bold",
                                     textColor=colors.white)

            story = []
            story.append(Paragraph("📋 Topic Database — Report", title_s))
            story.append(Paragraph(
                f"Generated: {datetime.today().strftime('%Y-%m-%d %H:%M')} | Topics: {len(data)}", sub_s))
            story.append(HRFlowable(width="100%", thickness=2, color=MED_BLUE, spaceAfter=14))

            # KPI table
            story.append(Paragraph("Summary KPIs", h2_s))
            total_n  = len(data); closed_n = len(data[data["status"]=="Closed"])
            kpi_d = [[Paragraph(h, hdr_s) for h in ["Total","Open","In Progress","Blocked","Closed","Done %"]],
                     [str(total_n), str(len(data[data["status"]=="Open"])),
                      str(len(data[data["status"]=="In Progress"])),
                      str(len(data[data["status"]=="Blocked"])),
                      str(closed_n), f"{round(closed_n/total_n*100) if total_n else 0}%"]]
            kpi_t = Table(kpi_d)
            kpi_t.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0),DARK_BLUE),
                ("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("FONTSIZE",(0,0),(-1,-1),9),
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("ROWBACKGROUNDS",(0,1),(-1,1),[LT_BLUE]),
                ("FONTNAME",(0,1),(-1,1),"Helvetica-Bold"),
                ("FONTSIZE",(0,1),(-1,1),14),
                ("GRID",(0,0),(-1,-1),0.5,colors.lightgrey),
                ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
            ]))
            story.append(kpi_t); story.append(Spacer(1, 14))

            # Category table
            story.append(Paragraph("Category Breakdown", h2_s))
            cat_d = [[Paragraph(h, hdr_s) for h in ["Category","Total","Open","In Progress","Blocked","Closed","% Done"]]]
            for cat in CATEGORIES:
                sub = data[data["category"]==cat]; t=len(sub); cl=len(sub[sub["status"]=="Closed"])
                cat_d.append([cat, str(t), str(len(sub[sub["status"]=="Open"])),
                               str(len(sub[sub["status"]=="In Progress"])),
                               str(len(sub[sub["status"]=="Blocked"])), str(cl),
                               f"{round(cl/t*100) if t else 0}%"])
            cat_ts = [("BACKGROUND",(0,0),(-1,0),DARK_BLUE),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                      ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),9),
                      ("ALIGN",(0,0),(-1,-1),"CENTER"),("GRID",(0,0),(-1,-1),0.5,colors.lightgrey),
                      ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6)]
            for ri, cat in enumerate(CATEGORIES, 1):
                cat_ts.append(("BACKGROUND",(0,ri),(-1,ri),CAT_C.get(cat,LT_BLUE)))
            cat_t = Table(cat_d); cat_t.setStyle(TableStyle(cat_ts))
            story.append(cat_t); story.append(PageBreak())

            # Detail table
            story.append(Paragraph("Topic Documentation", h2_s))
            col_ws = [1.0,3.5,4.0,2.0,2.0,2.0,2.2,5.0,5.0]
            col_ws_cm = [w*cm for w in col_ws]
            det_hdr = ["ID","Topic Group","Sub-Topic","Category","Date","PIC","Status",
                       "Problem Description","Root Cause / Corrective Actions"]
            det_d = [[Paragraph(h, hdr_s) for h in det_hdr]]
            data_s = data.sort_values("id").copy()
            data_s["_bucket"] = data_s["opening_date"].apply(aging_bucket)
            for _, r_ in data_s.iterrows():
                od = r_["opening_date"]
                prob  = (r_.get("problem_description","") or "")[:180]
                rca_ca= ((r_.get("root_cause_analysis","") or "")+" / "+(r_.get("corrective_actions","") or ""))[:180]
                det_d.append([
                    Paragraph(str(int(r_["id"])), cell_s),
                    Paragraph(r_["topic_group"], cell_s),
                    Paragraph(r_["sub_topic"], cell_s),
                    Paragraph(r_["category"], cell_s),
                    Paragraph(str(od.date()) if not pd.isna(od) else "—", cell_s),
                    Paragraph(r_.get("pic",""), cell_s),
                    Paragraph(r_["status"], cell_s),
                    Paragraph(prob, cell_s),
                    Paragraph(rca_ca, cell_s),
                ])
            det_ts = [
                ("BACKGROUND",(0,0),(-1,0),DARK_BLUE),("TEXTCOLOR",(0,0),(-1,0),colors.white),
                ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("FONTSIZE",(0,0),(-1,-1),7.5),
                ("VALIGN",(0,0),(-1,-1),"TOP"),("GRID",(0,0),(-1,-1),0.3,colors.lightgrey),
                ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F9FAFB")]),
            ]
            for ri, (_, r_) in enumerate(data_s.iterrows(), 1):
                det_ts.append(("BACKGROUND",(3,ri),(3,ri),CAT_C.get(r_["category"],colors.white)))
                det_ts.append(("BACKGROUND",(6,ri),(6,ri),STAT_C.get(r_["status"],colors.white)))
            det_t = Table(det_d, colWidths=col_ws_cm, repeatRows=1)
            det_t.setStyle(TableStyle(det_ts))
            story.append(det_t)

            doc.build(story)
            return buf.getvalue()

        except ImportError:
            return None

    try:
        import reportlab
        pdf_b = build_pdf(export_df)
        if pdf_b:
            st.download_button("⬇️ Download PDF Report", data=pdf_b,
                               file_name=f"topic_report_{datetime.today().strftime('%Y%m%d')}.pdf",
                               mime="application/pdf", type="primary")
        else:
            st.warning("PDF generation returned empty. Check logs.")
    except ImportError:
        st.warning("ReportLab not installed. Run: `pip install reportlab`")
        st.code("pip install reportlab")
