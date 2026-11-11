# Implementation Guide: Daily Gross Add – First and Second Recharge Cohort Analysis

## 1. Project Background

We need to develop an automated reporting solution to analyse the recharge behaviour of newly activated SIM cards (Daily Gross Adds / KG).

The main business objective is to understand how quickly newly activated customers perform their first and second qualifying recharges after activation.

The report must analyse:
- First recharge conversion and timing.
- Second recharge conversion and timing.
- Daily recharge distribution from Day 0 through Day 60.
- National-level performance.
- Regional-level performance.
- Weekly cohort trends and comparisons.
- Progress against predefined recharge targets.

**Important:** First and second recharge data are both available in the same Oracle source table. Do not assume that separate source tables are required.

## 2. Oracle Database Source

Use the following table as the primary data source:

```sql
SELECT *
FROM BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT;
```

Before implementing the reporting logic:

1. Inspect the complete table structure using Oracle data dictionary views.
2. Identify all column names, data types, and available dimensions.
3. Retrieve a representative sample of records.
4. Identify how first recharge and second recharge information are distinguished.
5. Identify the columns representing activation date, cohort week, KG base, recharge day, recharge counts, recharge percentages, and region.
6. Determine whether the source already contains calculated milestone percentages or whether these must be calculated in Python/SQL.
7. Check whether national and regional records are stored in the same table.
8. Identify any existing aggregation level to avoid double counting.

Do not invent column names or assume the table structure. Build the SQL based on the actual metadata and sample records.

## 3. Business Definitions

### 3.1 Daily Gross Add (KG)

KG represents newly activated SIM cards.

Each activation date represents a cohort. The KG Base is the total number of newly activated SIM cards in that cohort.

All recharge conversion calculations must use the appropriate cohort denominator.

### 3.2 First Recharge

First recharge represents the first qualifying recharge made by a newly activated subscriber.

The report must:
- Identify each subscriber's first qualifying recharge.
- Exclude recharge transactions of exactly 4 Birr.
- Track the first recharge from Day 0 through Day 60.
- Calculate the number of subscribers completing their first recharge on each day.
- Calculate cumulative first-recharge conversion at the required milestone days.

### 3.3 Second Recharge

Second recharge represents the second qualifying recharge made by a subscriber after activation.

The report must:
- Identify the second qualifying recharge chronologically.
- Exclude recharge transactions of exactly 4 Birr.
- Track the second recharge from Day 0 through Day 60.
- Calculate the number of subscribers completing their second recharge on each day.
- Calculate cumulative second-recharge conversion at the required milestone days.
- Produce the additional timing distribution shown in the sample report.

The implementation must ensure that a subscriber is counted only once for the first-recharge event and only once for the second-recharge event.

The treatment of multiple qualifying recharges on the same day must follow the actual transaction ordering and available source fields.

## 4. National-Level Reporting

Create two main reporting sections:

### 4.1 First Recharge – National Summary

The first-recharge summary must have the following structure:

| Cohort Date | Week | 3 Days | 4 Days | 7 Days | 10 Days | 16 Days | 20 Days |
|---|---|---:|---:|---:|---:|---:|---:|
| 29-Aug | WK5 | 24.0% | 28.2% | 36.6% | 41.8% | 49.2% | 52.4% |
| 5-Sep | WK6 | 23.2% | 27.3% | 35.3% | 40.6% | 48.6% | 51.7% |
| 12-Sep | WK7 | 23.7% | 28.1% | 36.5% | 42.0% | 49.3% | 52.5% |
| 19-Sep | WK8 | 23.1% | 27.2% | 35.3% | 40.4% | 48.4% | 51.8% |

**First-recharge targets:**

| Milestone | Target |
|---|---:|
| 3 Days | 30% |
| 4 Days | 40% |
| 7 Days | 50% |
| 10 Days | 60% |
| 16 Days | 70% |
| 20 Days | 90% |

Include:
- Weekly cohort performance.
- Four-week average.
- Target comparison.
- Historical weekly trends.
- Latest available cohort performance.

The four-week average must be calculated from the four relevant cohort records using the agreed business calculation. Do not simply average already rounded display values if the underlying counts and denominators are available.

### 4.2 Second Recharge – National Summary

Use the same cohort structure as the first-recharge summary.

**Second-recharge targets:**

| Milestone | Target |
|---|---:|
| 3 Days | 15% |
| 4 Days | 20% |
| 7 Days | 35% |
| 10 Days | 40% |
| 16 Days | 50% |
| 20 Days | 70% |

Include:
- Weekly cohort performance.
- Four-week average.
- Week-on-week comparisons.
- Target comparison.
- Historical weekly trends.
- Latest available cohort performance.

The summary must support comparisons such as:

- WK6 vs WK5
- WK7 vs WK6
- WK8 vs WK7
- WK9 vs WK8
- WK10 vs WK9
- WK11 vs WK10
- WK12 vs WK11

Calculate changes in percentage points, not relative percentage growth.

For example:

`WK6 vs WK5 = WK6 conversion rate - WK5 conversion rate`

A result of 0.5 percentage points should be displayed as `+0.5%` or `+0.5 pp`, consistently across the report.

## 5. National-Level Daily Detail Sheets

Create separate detail sheets for:

1. First Recharge Detail
2. Second Recharge Detail

Both sheets must follow the supplied sample layout.

### Required columns

| Column | Description |
|---|---|
| Week | Cohort week |
| Day | Day of the week for the activation date |
| Activation Date | Date the SIM was activated |
| KG Base | Total newly activated SIM cards |
| 3 Days | Cumulative milestone conversion |
| 4 Days | Cumulative milestone conversion |
| 7 Days | Cumulative milestone conversion |
| 10 Days | Cumulative milestone conversion |
| 16 Days | Cumulative milestone conversion |
| 20 Days | Cumulative milestone conversion |
| Day 0–Day 60 | Daily qualifying recharge counts |

For the second-recharge detail sheet, include the additional timing distribution columns:

| Column | Description |
|---|---|
| Total | Total second-recharge conversion |
| Day 1 | Second recharge on Day 1 |
| Day 2–5 | Second recharge during Days 2 to 5 |
| >5 to 10 | Second recharge after Day 5 through Day 10 |
| >10 to 15 | Second recharge after Day 10 through Day 15 |
| >15 | Second recharge after Day 15 |

The exact boundaries and whether Day 0 is included in these timing buckets must be validated against the existing sample calculations.

### Daily columns

Generate all daily columns from Day 0 through Day 60.

These columns must represent the daily number of unique subscribers completing the relevant recharge event, unless the source data confirms a different existing definition.

Do not confuse daily counts with cumulative counts.

## 6. Regional-Level Reporting

In addition to the national report, implement regional reporting.

The regional analysis must support the following regions:

1. East Addis
2. West Addis
3. Central
4. South
5. West
6. East 1
7. East 2
8. North West
9. North East
10. North

Use the actual region field and values available in the Oracle source table. Validate the mapping and standardise region names before generating reports.

### 6.1 Regional Second-Recharge Summary

Create a consolidated regional summary sheet.

The layout must follow the supplied regional sample.

Each region must have its own milestone columns:

- 3 Days
- 4 Days
- 7 Days
- 10 Days
- 16 Days
- 20 Days

The report should be structured as follows:

| Cohort Date | Week | East Addis | West Addis | Central | South | West | East 1 | East 2 | North West | North East | North |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 29-Aug | WK5 | Milestone rates | Milestone rates | Milestone rates | ... | ... | ... | ... | ... | ... | ... |
| 5-Sep | WK6 | Milestone rates | Milestone rates | Milestone rates | ... | ... | ... | ... | ... | ... | ... |

This is a conceptual layout. The actual Excel sheet should use grouped headers, with the six milestone columns nested under each region.

If a region does not have sufficient observation history for a milestone, leave that milestone blank rather than treating it as zero.

### 6.2 Regional Four-Week Average

Calculate the four-week average for each region and each milestone.

The calculation must use the same cohort and denominator logic as the national report.

### 6.3 Regional After-vs-Before Comparison

Include an `After vs Before` comparison row for each milestone and region.

Use the agreed four-week comparison periods.

The calculation should be:

`After vs Before = After-period conversion rate - Before-period conversion rate`

Clearly document which four weeks are considered the before period and which four weeks are considered the after period.

### 6.4 Regional Week-on-Week Comparison

Include weekly comparisons for each region:

- WK6 vs WK5
- WK7 vs WK6
- WK8 vs WK7
- WK9 vs WK8
- WK10 vs WK9
- WK11 vs WK10
- WK12 vs WK11

Calculate differences in percentage points.

### 6.5 Regional Target Comparison

Include a target comparison row for each region.

For the second-recharge report, the 20-day target is 70%.

The comparison should show:

`Actual conversion rate - Target`

For example, an actual conversion of 39.5% against a 70% target gives `-30.5 percentage points`.

Keep the target comparison separate from week-on-week and four-week comparisons.

## 7. Regional Detail Sheets

Create an individual second-recharge detail sheet for each region.

For example:

- Second Recharge - East Addis
- Second Recharge - West Addis
- Second Recharge - Central
- Second Recharge - South
- Second Recharge - West
- Second Recharge - East 1
- Second Recharge - East 2
- Second Recharge - North West
- Second Recharge - North East
- Second Recharge - North

Each regional detail sheet must follow the same structure as the national second-recharge detail sheet:

- Week
- Day
- Activation Date
- Regional KG Base
- 3 Days
- 4 Days
- 7 Days
- 10 Days
- 16 Days
- 20 Days
- Day 0 through Day 60
- Total
- Day 1
- Day 2–5
- >5 to 10
- >10 to 15
- >15

All counts and percentages must be calculated using subscribers belonging to that region and the corresponding regional KG Base.

Do not use the national KG Base as the denominator for regional conversion calculations.

If regional first-recharge reporting is required, use the same regional structure and first-recharge target definitions. Keep its calculations separate from second recharge.

## 8. Formula Relationships and Workbook Design

The national and regional summary sheets must be connected to their respective detail sheets.

The preferred design is:

1. Detail sheets contain the underlying cohort-level counts and calculated milestone values.
2. National summary sheets reference the national detail sheets.
3. Regional summary sheets reference the relevant regional detail sheets.
4. Four-week averages and weekly comparisons are calculated from the underlying cohort data.
5. Target comparisons reference the defined target values.

Where the existing workbook uses Excel formulas, preserve the formula-driven relationships rather than hardcoding summary values.

If the report is generated using Python and openpyxl:
- Write the detail data first.
- Add formulas to summary sheets where practical.
- Use consistent cell references and formula ranges.
- Avoid hardcoded values that should be calculated from detail sheets.
- Ensure formulas remain valid when new weekly cohorts are added.
- Account for the fact that openpyxl writes formulas but does not calculate them. The workbook should be recalculated by a compatible spreadsheet application or use precomputed values where formula evaluation is unavailable.

## 9. Data Validation and Calculation Rules

Before finalising the report, validate the following:

1. Each activation cohort has the correct KG Base.
2. A subscriber is counted only once for each recharge milestone.
3. First recharge is not confused with second recharge.
4. Transactions of exactly 4 Birr are excluded.
5. Daily recharge counts reconcile with the underlying qualifying subscriber records.
6. Milestone percentages reconcile with the cumulative daily counts.
7. Regional KG Base totals reconcile with the national KG Base, where the regional mapping is complete and mutually exclusive.
8. Regional recharge counts reconcile with national counts where the source data supports that comparison.
9. Four-week averages use the correct four cohorts.
10. Week-on-week comparisons use the correct consecutive cohorts.
11. Incomplete cohorts are not incorrectly reported as zero.
12. Missing observation days are distinguished from genuine zero-recharge days.
13. No duplicate subscribers are introduced through joins or aggregation.
14. Percentage calculations use unrounded counts and denominators.

### Important validation

The sample reports contain some irregular-looking values and incomplete milestone columns for recent cohorts. Check the underlying source data and formulas before reproducing these values.

Do not automatically interpret missing values as zero.

Do not assume that all cohorts have completed 20 or 60 days of observation.

## 10. Technical Implementation Requirements

First inspect the existing project and identify its current architecture.

If the project uses Python, prefer the existing dependencies and conventions. Do not introduce a new framework without a clear reason.

The implementation should:

1. Connect to Oracle using the project's existing database connection configuration.
2. Query `BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT`.
3. Extract the required national and regional data.
4. Validate and normalise the data.
5. Generate national first-recharge detail and summary sheets.
6. Generate national second-recharge detail and summary sheets.
7. Generate the consolidated regional summary.
8. Generate individual regional second-recharge detail sheets.
9. Apply Excel formatting and formula relationships.
10. Save the final report to the existing configured output directory.

Use parameterised date filters where appropriate. Avoid unnecessarily querying the full historical table if the report only requires a defined reporting period.

Do not change existing working database connection logic, project configuration, or unrelated reporting functionality.

## 11. Excel Formatting Requirements

The workbook should be professional and suitable for management reporting.

Apply:
- Clear and consistent sheet names.
- Grouped regional headers.
- Distinct formatting for milestone columns and daily columns.
- Percentage formatting for conversion rates.
- Number formatting with thousand separators for subscriber counts.
- Consistent date formatting.
- Freeze panes for the cohort identification columns.
- Filters where useful.
- Appropriate column widths.
- Conditional formatting to highlight performance against targets.
- Clear separation of national and regional reporting sections.

Avoid excessive decoration. Prioritise readability and consistency across all sheets.

## 12. Expected Final Workbook Structure

The final workbook should contain the following sheets:

| Sheet | Purpose |
|---|---|
| 1st Recharge Summary | National first-recharge milestone summary |
| 1st Recharge Detail | National first-recharge daily cohort detail |
| 2nd Recharge Summary | National second-recharge milestone summary |
| 2nd Recharge Detail | National second-recharge daily cohort detail |
| Regional 2nd Recharge Summary | Consolidated regional milestone comparison |
| 2nd Recharge - East Addis | Regional second-recharge detail |
| 2nd Recharge - West Addis | Regional second-recharge detail |
| 2nd Recharge - Central | Regional second-recharge detail |
| 2nd Recharge - South | Regional second-recharge detail |
| 2nd Recharge - West | Regional second-recharge detail |
| 2nd Recharge - East 1 | Regional second-recharge detail |
| 2nd Recharge - East 2 | Regional second-recharge detail |
| 2nd Recharge - North West | Regional second-recharge detail |
| 2nd Recharge - North East | Regional second-recharge detail |
| 2nd Recharge - North | Regional second-recharge detail |

If regional first-recharge reporting is confirmed as part of the scope, add the corresponding regional summary and detail sheets without mixing their calculations with second recharge.

## 13. Implementation Approach

Follow these phases.

**Phase 1 – Source inspection**
- Inspect the Oracle table structure.
- Retrieve sample records.
- Identify the first- and second-recharge fields.
- Identify region and cohort dimensions.
- Document the source-to-report mapping.

**Phase 2 – Business logic**
- Confirm the recharge event definitions.
- Confirm Day 0 and Day 1 indexing.
- Confirm the milestone boundaries.
- Confirm the timing bucket definitions.
- Confirm the denominator and incomplete-cohort treatment.

**Phase 3 – Data extraction and calculations**
- Build the Oracle query.
- Validate national cohort calculations.
- Validate regional cohort calculations.
- Reconcile daily counts and milestone rates.

**Phase 4 – Workbook generation**
- Generate the detail sheets.
- Generate national summaries.
- Generate regional summaries.
- Add formula relationships.
- Apply formatting.

**Phase 5 – Validation**
- Compare the generated workbook with the supplied samples.
- Check cohort totals and regional reconciliation.
- Verify the weekly and four-week calculations.
- Confirm that incomplete cohorts are handled correctly.

## 14. First Deliverable Before Coding

Before modifying the existing project, provide:

1. The actual Oracle table structure.
2. A mapping of source columns to report fields.
3. The proposed first-recharge calculation logic.
4. The proposed second-recharge calculation logic.
5. The proposed regional aggregation logic.
6. The proposed Excel sheet structure.
7. Any business rules that cannot be established from the source data.

Only proceed with implementation after the source structure and calculation logic are understood.

**Final objective:** Produce an automated, accurate, formula-connected national and regional first- and second-recharge cohort reporting workbook using `BETELHEM_GIRMA.TBL_DAILY_TOTAL_SECOND_RECHARGE_COHORT_RPT` as the primary Oracle source.