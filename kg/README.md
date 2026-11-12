# Daily Gross Add (KG) — First and Second Recharge Cohort Reporter

An end-to-end automated reporting and executive communication pipeline for Safaricom Ethiopia. It tracks the first and second recharge conversion velocities of newly activated SIM cards (Daily Gross Adds / Kindergarten - KG), generates a comprehensive 14-tab formula-linked Excel workbook, and delivers a branded executive HTML email.

---

## 1. Architecture Overview

Modeled after the high-performance design of `bundle-performance`, this project provides:
1. **Multi-Source Database Abstraction**:
   - **Local DuckDB (`data/dummy_database.duckdb`)**: Pre-seeded with 102 national daily cohorts and 810 regional daily cohort records extracted directly from the historical baseline workbook.
   - **Production Oracle EDW (`python-oracledb`)**: Queries `BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT` and `TBL_DAILY_SECOND_RECHARGE_COHORT_RPT`.
   - **Zero-Downtime Fallback**: If Oracle credentials are missing or the EDW network is unreachable, it automatically and gracefully falls back to DuckDB.
2. **Formula-Driven 14-Tab Excel Workbook**:
   - `1RC-Wk`: National 1st Recharge Weekly Summary with target variances and `=AVERAGEIFS` formulas.
   - `1RC`: National 1st Recharge Daily Cohort Detail (Day 0 through Day 60).
   - `2RC-Wk`: National 2nd Recharge Weekly Summary with 4-week averages and Week-on-Week differences.
   - `2RC`: National 2nd Recharge Daily Cohort Detail (Day 0 through Day 60 + Timing buckets).
   - `2RC-Wk (Divison)`: Consolidated regional milestone comparison across all 10 regions side-by-side.
   - 10 Individual Regional Detail Sheets (`2RC (East ADDIS)`, `2RC (West ADDIS)`, `2RC (Central)`, `2RC (South)`, `2RC (West)`, `2RC (East 1)`, `2RC (East 2)`, `2RC (North West)`, `2RC (NorthEast)`, `2RC (North)`).
3. **Executive Visual Charting**:
   - High-resolution 2-panel chart rendering national weekly milestone trajectories and regional 20-day conversion rankings vs the 70% target.
4. **Responsive Safaricom Branded HTML Email**:
   - MSO Outlook conditional layout (`<!--[if mso]>`).
   - Executive KPI cards, milestone delta pills, regional leaderboard, and embedded visual charts.
   - Dual image rendering: Base64 data-URI for standalone browser previews and CID inline for SMTP Outlook delivery.

---

## 2. Directory Structure

```
kg/
├── config/
│   ├── __init__.py
│   └── settings.py              # Configuration, milestones, canonical regions & paths
├── data/
│   └── dummy_database.duckdb    # Local DuckDB database pre-seeded with sample cohorts
├── sql/
│   ├── national_cohort.sql      # National cohort extraction query
│   └── regional_cohort.sql      # Regional cohort extraction with CASE region ordering
├── src/
│   └── kg/
│       ├── __init__.py
│       ├── assets/
│       │   └── safaricom_email_footer.png  # Official Safaricom email footer banner
│       └── services/
│           ├── __init__.py
│           ├── dummy_db.py         # DuckDB seeder from baseline Excel workbook
│           ├── database.py         # Database abstraction (DuckDB/Oracle with fallback)
│           ├── excel_formatting.py # Safaricom colors, fonts, borders & formatters
│           ├── excel_report.py     # 14-sheet Excel generator with dynamic formulas
│           ├── visual_charts.py    # Matplotlib executive graphical chart generator
│           ├── email_template.py   # Responsive Safaricom executive HTML email
│           ├── email_service.py    # SMTP client & standalone preview generator
│           └── runner.py           # Pipeline orchestrator
├── reports/
│   ├── excel/                   # Generated KG_Recharge_Tracker_*.xlsx workbooks
│   ├── previews/                # Desktop & mobile HTML email preview files
│   └── charts/                  # Generated PNG performance charts
├── send_kg_email.py             # CLI entry point
├── daily_kg_email.bat           # Windows Task Scheduler automation batch script
├── pyproject.toml               # Python dependencies
└── .env                         # Local environment settings (DB and SMTP)
```

---

## 3. Quick Start

### Installation

Install dependencies using `uv`:

```powershell
cd "c:\Users\Mekin.Jemal\OneDrive - Safaricom Ethiopia\Desktop\Automation\kg"
uv sync
```

### Running with DuckDB (Offline / Development Mode)

Generate the 14-sheet Excel report, visual chart, and standalone browser previews:

```powershell
uv run python send_kg_email.py --source duckdb --preview-only
```

### Running with Oracle (Production EDW Mode)

```powershell
uv run python send_kg_email.py --source oracle --preview-only
```
*(If Oracle credentials are not set, it will safely log a warning and fall back to DuckDB).*

### Dispatching Email via SMTP

```powershell
uv run python send_kg_email.py --source duckdb --send
```

---

## 4. SQL Queries

### National Cohort Query (`sql/national_cohort.sql`)
```sql
SELECT *
FROM BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT
ORDER BY ACTIVATION_DATE;
```

### Regional Cohort Query (`sql/regional_cohort.sql`)
```sql
SELECT
    CASE
        WHEN L.DS_EXEC_REGION = 'West Addis'  THEN 'A.West Addis'
        WHEN L.DS_EXEC_REGION = 'East Addis'  THEN 'B.East Addis'
        WHEN L.DS_EXEC_REGION = 'Central'     THEN 'C.Central'
        WHEN L.DS_EXEC_REGION = 'South'       THEN 'D.South'
        WHEN L.DS_EXEC_REGION = 'North West'  THEN 'E.North West'
        WHEN L.DS_EXEC_REGION = 'North East'  THEN 'F.North East'
        WHEN L.DS_EXEC_REGION = 'East 1'      THEN 'G.East 1'
        WHEN L.DS_EXEC_REGION = 'East 2'      THEN 'H.East 2'
        WHEN L.DS_EXEC_REGION = 'West'        THEN 'I.West'
        WHEN L.DS_EXEC_REGION = 'North'       THEN 'J.North'
        WHEN L.DS_EXEC_REGION = 'Afar'        THEN 'K.Afar'
        ELSE 'L.Unknown'
    END AS REGION,
    L.*
FROM BETELHEM_GIRMA.TBL_DAILY_SECOND_RECHARGE_COHORT_RPT L
WHERE L.DS_EXEC_REGION IN (
    'East Addis',
    'West Addis',
    'Central',
    'South',
    'West',
    'East 1',
    'East 2',
    'North West',
    'North East',
    'North'
)
AND L.REPORT_TYPE = '2nd Purchase'
ORDER BY REGION, ACTIVATION_DATE;
```

---

## 5. Milestone Targets

| Milestone | 1st Recharge Target | 2nd Recharge Target |
|:---|:---:|:---:|
| **3 Days** | 30% | 15% |
| **4 Days** | 40% | 20% |
| **7 Days** | 50% | 35% |
| **10 Days** | 60% | 40% |
| **16 Days** | 70% | 50% |
| **20 Days** | 90% | 70% |

---

## 6. Project Documentation Suite

Comprehensive guides matching the architecture of `bundle-performance/docs` are located in the `docs/` folder:

1. **[Quick Project & Operations Guide](file:///c:/Users/Mekin.Jemal/OneDrive%20-%20Safaricom%20Ethiopia/Desktop/Automation/kg/docs/PROJECT_GUIDE.md)** (`docs/PROJECT_GUIDE.md`):
   - Fast operational guide covering setup (`uv sync`), `.env` configuration, CLI commands, output paths, and troubleshooting.
2. **[End-to-End System Architecture & Operations Manual](file:///c:/Users/Mekin.Jemal/OneDrive%20-%20Safaricom%20Ethiopia/Desktop/Automation/kg/docs/END_TO_END_PROJECT_GUIDE.md)** (`docs/END_TO_END_PROJECT_GUIDE.md`):
   - In-depth technical architecture, data normalization rules, 15-sheet Excel formula breakdown, MSO email mechanics, and failover mechanics.
3. **[Enterprise Excel Automation AI Agent Blueprint](file:///c:/Users/Mekin.Jemal/OneDrive%20-%20Safaricom%20Ethiopia/Desktop/Automation/kg/docs/EXCEL_AUTOMATION_AI_AGENT_BLUEPRINT.md)** (`docs/EXCEL_AUTOMATION_AI_AGENT_BLUEPRINT.md`):
   - Architectural framework for transforming complex SQL queries and multi-tab Excel workbooks into autonomous production pipelines.


