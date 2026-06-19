# Tenant Renewal Risk Scorecard Project

This project converts the original Lease Renewal Probability Model into a tenant renewal risk and credit-style scorecard.

The business question is no longer only "Will the tenant renew?". The model now answers: "Should management offer this tenant a renewal?"

## Folder structure

- `data/` stores the updated Excel masterfile.
- `database/` stores the SQLite database.
- `sql/` stores the database table schema.
- `python/` stores the import and reporting scripts.
- `excel/` stores the final management report.

## First-time setup on Mac

From Terminal:

```bash
cd ~/Projects/Lease_Renewal_Probability_Model
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Step 1: Import Excel into SQLite

Place `Lease_Masterfile_Risk_Model.xlsx` in the `data/` folder.

Run:

```bash
python python/import_excel_to_sqlite.py
```

This creates or replaces the `tenant_risk_scorecard` table in `database/lease_renewal.db`.

## Step 2: Export the upgraded management report

Run:

```bash
python python/export_management_report_upgraded.py
```

This creates `excel/Tenant_Renewal_Management_Report_Upgraded.xlsx`.

## Current scorecard logic

The current tenant risk score is based on:

- Payment score overall
- Basic rental payment score
- Utilities payment score
- Arrears incidents
- Lease longevity
- Deposit coverage

Recommended interpretation:

- A = Strongly Renew
- B = Renew
- C = Renew with Review
- D = Management Review
- F = Do Not Renew

This is a deterministic management scorecard, not a machine-learning model yet. Once sufficient historical outcomes exist, it can be upgraded to logistic regression.
