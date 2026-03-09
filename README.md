# Knowledge-Database

An automated topic tracking and reporting tool for **Elements**, **Assembly**, and **Cross** category issues. Built with Python and Streamlit.

---

Quality Topics Tracker – Streamlit Application
A smart, automated system for tracking Quality topics across Elements, Assembly, and Cross-functional areas.
Designed for Quality Assurance Management, IATF compliance, and Lean Six Sigma workflows.

# Overview
The Quality Topics Tracker is a Streamlit-based web application backed by an Excel data store.
It provides an automated, structured, and audit-ready platform to:

Register and document Quality issues/topics
Automatically calculate aging, risk scoring, and escalation
Visualize category breakdowns, trends, and Pareto charts
Upload supporting evidence (pictures, files)
Export topics for QAMM/GQA management reporting

The system is fully modular and can be extended or migrated to an SQL backend without UI changes.

# Features
1. Topic Database
Each topic includes:

Topic group & sub-topic
Category (Elements / Assembly / Cross)
PIC
Problem description
Root cause analysis
Corrective actions
Customer impact flag
Status lifecycle
Opening date and milestone dates
Links to 8Ds, meeting minutes, shared folders
Attachments (photos, documents)

The system automatically adds:

TopicID (unique)
DaysOpen
Aging bucket classification
Risk scoring
Escalation flag


2. Dashboards & Analytics
Home Dashboard

KPI cards: Total / Open / Closed / Blocked / Completion %
Aging bucket bar chart
Category Pareto chart
High-risk & overdue agenda list

Analytics Page

Monthly opening trend
Root cause Pareto (top 20 issues)
Cross-category comparisons


3. Smart Topic Creation Form
The input form includes:

Preconfigured categories, severities, and statuses
Multi-file upload
Auto-generated dates and default values
Easy link insertion
Automatic naming & saving of attachments


4. Data Storage
Default Storage

Excel file: data/topics.xlsx
Attachments stored under: data/attachments/
Automatic backups: data/backups/

Optional
Switch to SQLite by enabling the flag in config.py.

5. One‑Click Export
The Exports page provides:

Excel export
CSV export
PPTX export (slides summarizing each topic)

Useful for:

QAMM reports
Customer meetings
Management updates


📁 Project Structure
qa_topics_app/
│
├─ app.py
├─ pages/
│   ├─ 1_New_Topic.py
│   ├─ 2_Topic_Explorer.py
│   ├─ 3_Analytics.py
│   ├─ 4_Exports.py
│
├─ components/
│   ├─ kpi_cards.py
│   ├─ charts.py
│   ├─ forms.py
│   └─ tables.py
│
├─ core/
│   ├─ config.py
│   ├─ schema.py
│   ├─ logic.py
│   └─ services.py
│
├─ io/
│   ├─ repo_excel.py
│   ├─ repo_sqlite.py
│   ├─ migration.py
│   └─ exporters.py
│
├─ utils/
│   ├─ cache.py
│   ├─ dates.py
│   ├─ ids.py
│   └─ ui.py
│
├─ data/
│   ├─ topics.xlsx
│   ├─ attachments/
│   └─ backups/
│
├─ requirements.txt
└─ README.md


# Installation
1. Clone the repository
Shellgit clone https://github.com/<your-repo>/qa_topics_app.gitcd qa_topics_appWeitere Zeilen anzeigen
2. Install dependencies
Shellpip install -r requirements.txtWeitere Zeilen anzeigen

▶️ Running the Application
Shellstreamlit run app.pyWeitere Zeilen anzeigen
The application will open automatically in your browser at:
http://localhost:8501

⚙️ Configuration
You can adjust categories, statuses, and features in:
core/config.py

Examples:

Enable/disable risk scoring
Switch from Excel → SQLite
Adjust aging buckets
Add new categories


🧠 Risk Scoring Logic
The risk score is calculated from:
Severity (1–3)
+ Customer impact (0 or +2)
+ Aging score (0–3)
= RiskScore (0–8)

High-risk topics or topics open >180 days trigger escalation.

# Data Model
Each topic contains canonical fields such as:
TopicID
TopicGroup
SubTopic
Category
PIC
CustomerImpact
Severity
ProblemDescription
RootCauseAnalysis
CorrectiveActions
Status
NextSteps
OpeningDate
FirstResponseDate
RCADoneDate
CADoneDate
ClosedDate
DaysOpen
AgingBucket
RiskScore
EscalationFlag
Link8D
LinkMinutes
LinkFolder
AttachmentPaths


# Charts & Visualizations
Included charts:

Aging bucket distribution
Category Pareto
Root cause Pareto
Monthly trend
Agenda candidates

All charts automatically update based on filters.

# Exports
Excel
Pythonexport_excel(df, "export_topics.xlsx")Weitere Zeilen anzeigen
CSV
Pythonexport_csv(df, "export_topics.csv")Weitere Zeilen anzeigen
PPTX
Generates slides with:

Topic ID
Category
PIC
Status
Problem summary
RCA summary
CA summary
Next steps
Risk score & aging


# Data Privacy
No external services are used.
All data remains local within the data/ folder.
For corporate use, ensure compliance with internal rules on storage and attachments.
