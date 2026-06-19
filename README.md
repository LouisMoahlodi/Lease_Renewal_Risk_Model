# Tenant Renewal Risk Scorecard Project

This project converts the original Lease Renewal Probability Model into a tenant renewal risk and credit-style scorecard.

The business question is no longer only "Will the tenant renew?". The model now answers: "Should management offer this tenant a renewal?"

## Folder structure

- `data/` stores the updated Excel masterfile.
- `database/` stores the SQLite database.
- `sql/` stores the database table schema.
- `python/` stores the import and reporting scripts.
- `excel/` stores the final management report.

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
