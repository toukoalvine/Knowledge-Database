# QM Topic Dashboard — Streamlit App

## Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the app
```bash
streamlit run app.py
```

The app opens at **http://localhost:8501**

## Usage

- **Upload your Excel** via the sidebar (must contain a sheet named `Topic Database` with headers in row 2)
- **Filter** by Category, Severity, Status, Escalated, PIC, Days Open
- **Search** any text across all fields
- **Sort** by ID, Days Open, Severity, Status, or Category
- **Expand** any topic row to see full details (problem, root cause, corrective actions, next steps)
- **Export** the currently filtered view to Excel

## Data columns expected
ID, Topic Group, Sub-Topic, Category, Severity, Opening Date, Close Date,
PIC NED, PIC HQ, Status, Cust. Impact, Days Open, Aging Bucket, Escalated,
Problem Description, Root Cause Analysis, Corrective Actions,
Prevention of recurrence, Next Steps, Milestones / Dates

## Note on Risk Score
Risk Score column is intentionally excluded from the dashboard view.
