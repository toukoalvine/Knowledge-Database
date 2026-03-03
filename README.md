# Knowledge-Database

An automated topic tracking and reporting tool for **Elements**, **Assembly**, and **Cross** category issues. Built with Python and Streamlit.

---

## Features

- **Dashboard** — Live KPIs, aging buckets, category breakdown, and completion progress
- **Topic List** — Browse, filter, and edit all topics in one place
- **New Topic Form** — Log issues with full documentation fields
- **Export** — Generate reports as Excel (`.xlsx`) or PDF at any time

---

## Requirements

- Python 3.10 or higher
- The following packages (see `requirements.txt`):

| Package | Purpose |
|---|---|
| `streamlit` | Web app framework |
| `pandas` | Data handling |
| `openpyxl` | Excel export |
| `reportlab` | PDF export |

---

## Installation & Setup

**1. Clone or download the project files**

Make sure you have these files in the same folder:
```
topic_database_app.py
requirements.txt
README.md
```

**2. (Optional) Create a virtual environment**
```bash
python -m venv venv
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows
```

**3. Install dependencies**
```bash
pip install -r requirements.txt
```

**4. Run the app**
```bash
streamlit run topic_database_app.py
```

The app will open automatically in your browser at `http://localhost:8501`.

---

## Data Storage

All topics are saved locally to **`topics_data.json`** in the same folder as the app. This file is created automatically on first run and persists between sessions.

> ⚠️ Do not delete `topics_data.json` unless you want to reset all data. Back it up regularly.

To reset to the demo data, simply delete `topics_data.json` and restart the app.

---

## App Sections

### 📊 Dashboard
The main overview page. Displays:
- **KPI cards**: Total, Open, In Progress, Blocked, Closed, and overall Completion %
- **Progress bar**: Visual completion indicator
- **Aging buckets**: Count of non-closed topics grouped by time since opening (0–3 months, 3–6 months, 6–12 months, >1 year)
- **Category breakdown**: Per-category totals and status counts for Elements, Assembly, and Cross

### 📋 All Topics
Browse all topics as expandable cards. Each card shows the full topic record including problem description, root cause analysis, corrective actions, next steps, and customer impact.

Use the **sidebar filters** to narrow by Category, Status, or free-text search.

Click **✏️ Edit** on any topic to update its fields inline.

### ➕ New Topic
Form to log a new topic. Required fields are marked with `*`. Supported fields:

| Field | Description |
|---|---|
| Topic Group | High-level grouping (e.g. "Welding Defects") |
| Sub-Topic | Specific issue within the group |
| Category | Elements / Assembly / Cross |
| Opening Date | Date the issue was first identified |
| PIC | Person in Charge |
| Problem Description | What was observed |
| Root Cause Analysis | Known or suspected root causes |
| Corrective Actions | Actions taken so far |
| Status | Open / In Progress / Blocked / Closed |
| Next Steps | Planned actions |
| Customer Impact | Yes / No |
| Pictures | Optional image attachments (PNG, JPG) |

### 📤 Export
Generate a formatted report from the current dataset. A toggle lets you export either all topics or only the currently filtered set.

**Excel export** (`.xlsx`) — two sheets:
- *Summary Dashboard*: KPI table, aging buckets, and category breakdown
- *Topic Documentation*: Full topic table with color-coded Category and Status columns, frozen header row

**PDF export** (`.pdf`, landscape A4):
- KPI summary table
- Category breakdown table
- Full topic detail table with color coding

---

## Sidebar Filters

The sidebar is visible on all pages and applies to the **All Topics** view and the **Export** (when "ignore filters" is off):

- **Category** — filter by Elements, Assembly, Cross (multi-select)
- **Status** — filter by Open, In Progress, Blocked, Closed (multi-select)
- **Search** — free-text search across Topic Group, Sub-Topic, and Problem Description

---

## Color Coding

| Category | Color |
|---|---|
| Elements | 🔵 Blue |
| Assembly | 🟢 Green |
| Cross | 🟡 Yellow |

| Status | Color |
|---|---|
| Open | 🔴 Red |
| In Progress | 🔵 Blue |
| Blocked | 🟡 Yellow |
| Closed | 🟢 Green |

---

## Troubleshooting

**PDF export button not showing**
Install ReportLab: `pip install reportlab`

**App won't start**
Make sure all packages are installed: `pip install -r requirements.txt`

**Data not saving between sessions**
Check that the app has write permission in its folder. The `topics_data.json` file must be writeable.

**Port already in use**
Run on a different port: `streamlit run topic_database_app.py --server.port 8502`
