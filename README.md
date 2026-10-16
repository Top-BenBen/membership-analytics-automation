# Membership Analytics & Reporting Automation

A portfolio project demonstrating how Python can turn a structured membership snapshot into a cleaned dataset and a multi sheet Excel analytics report. The included records are synthetic and created for demonstration.

## Project overview

Membership records can contain inconsistent text, missing fields, duplicate rows, and dates stored in different formats. This project standardises key fields, calculates demographic and tenure measures, surfaces data quality checks, and generates a formatted Excel workbook for review.

## What it does

- Reads membership data from CSV or Excel.
- Checks that required columns are present.
- Standardises text fields and parses date columns.
- Removes exact duplicate rows and flags missing or repeated membership numbers for review.
- Calculates age, five year age cohorts, and membership tenure bands.
- Summarises membership by gender, relationship, company, industry, benefit option, age cohort, and tenure.
- Exports an Excel workbook with an executive summary, analysis sheets, data quality checks, cleaned records, and charts.

## Technology stack

- Python 3.10 or later
- pandas
- NumPy
- openpyxl
- Excel for viewing the generated report

## Repository structure

```text
membership-analytics-automation/
├── data/
│   └── sample/
│       └── sample_membership.csv
├── outputs/
│   └── membership_analytics_report.xlsx
├── src/
│   └── membership_analytics.py
├── .gitignore
├── LICENSE
├── README.md
└── requirements.txt
```

## Quick start on Windows

Open a terminal in the repository folder and run:

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python src/membership_analytics.py --as-of 2026-10-16
```

The script reads `data/sample/sample_membership.csv` by default and creates `outputs/membership_analytics_report.xlsx`. Change the `--as-of` date to match the reporting date for your snapshot.

To process a different CSV or Excel file, provide the input and output paths:

```powershell
python src/membership_analytics.py "data/my_membership.xlsx" "outputs/my_report.xlsx" --as-of 2026-10-16
```

## Required input columns

The input file must contain these columns:

- `MEMBERSHIP NO`
- `Date of Birth`
- `GENDER`
- `COMPANY`
- `INDUSTRY`
- `Registration Date`
- `Join Date`
- `BENEFIT OPTION`
- `MEMBER TYPE`
- `RELATIONSHIP TYPE`

The input is assumed to be an **active membership snapshot** as of the reporting date, not a complete historical register. Monthly or quarterly growth should only be calculated when comparable dated snapshots are available.

## Output workbook

The workbook includes these sheets:

1. Executive Summary
2. Demographics
3. By Company
4. By Industry
5. Plan Analysis
6. Age Analysis
7. Tenure
8. Data Quality
9. Cleaned Data

A sample workbook is included in `outputs/` so visitors can inspect the output without running the script first.

## Data privacy and responsible use

The included CSV contains synthetic records and does not represent real members or clients. Do not publish real member identifiers, dates of birth, client names, employer data, internal reports, credentials, or confidential business logic. Before adapting this project to workplace data, confirm that you are authorised to do so and follow your employer's data protection policies.

## Limitations and possible extensions

- Repeated membership numbers are flagged for investigation rather than automatically removed, because business rules determine whether they are valid.
- The report describes patterns and data quality; it does not establish causation or make eligibility decisions.
- The current version analyses one snapshot at a time. Trend analysis requires comparable historical snapshots.
- Possible extensions include automated tests, configurable column mapping, snapshot to snapshot trend analysis, and a lightweight interactive interface.

## Portfolio summary

This project demonstrates data cleaning, validation, feature engineering, descriptive analytics, and Excel report automation. It is a portfolio example, not a production certified insurance system.
