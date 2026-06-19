import sqlite3
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]

excel_path = BASE_DIR / "data" / "Lease_Masterfile_Risk_Model.xlsx"
db_path = BASE_DIR / "database" / "lease_renewal.db"
sql_path = BASE_DIR / "sql" / "create_tables.sql"

sheet_name = "TenantMasterList"

if not excel_path.exists():
    raise FileNotFoundError(f"Excel file not found: {excel_path}")

if not sql_path.exists():
    raise FileNotFoundError(f"SQL file not found: {sql_path}")

df = pd.read_excel(excel_path, sheet_name=sheet_name)

# Exact column mapping from the current masterfile to SQL-safe column names.
column_map = {
    "Tenant ID": "tenant_id",
    "PROPERTY": "property",
    "UNIT NO": "unit_no",
    "ERF NO": "erf_no",
    "TENANT NAME": "tenant_name",
    "BASIC RENT": "basic_rent",
    "DEPOSIT RENT": "deposit_rent",
    "DEPOSIT UTILITIES": "deposit_utilities",
    "KEY DEPOSIT": "key_deposit",
    "CURRENT LEASE START DATE": "current_lease_start_date",
    "CURRENT LEASE END DATE": "current_lease_end_date",
    "NOTICE DUE DATE": "notice_due_date",
    "ESCALATION DUE": "escalation_due",
    "LEASE TYPE \n(FIXED TERM/MONTH -TO-MONTH)": "lease_type",
    "RENEWAL STATUS ": "renewal_status",
    "Comments": "comments",
    "Initial Lease Commencement Date": "initial_lease_commencement_date",
    "NEXT ESCALATION RENTAL": "next_escalation_rental",
    "PAYMENT SCORE (OVERALL)": "payment_score_overall",
    "PAYMENT SCORE - BASIC RENTAL": "payment_score_basic_rental",
    "PAYMENT SCORE - UTILITIES": "payment_score_utilities",
    "TENURE MONTHS": "tenure_months",
    "ARREARS INCIDENTS (CALC)": "arrears_incidents_calc",
    "LEASE LONGEVITY MONTHS": "lease_longevity_months",
    "DEPOSIT COVERAGE SCORE": "deposit_coverage_score",
    "TENANT RISK SCORE": "tenant_risk_score",
    "TENANT GRADE": "tenant_grade",
    "RENEWAL RECOMMENDATION": "renewal_recommendation",
}

missing = [col for col in column_map if col not in df.columns]
if missing:
    raise ValueError("Missing required Excel columns: " + ", ".join(missing))

df = df.rename(columns=column_map)
required_columns = list(column_map.values())
df = df[required_columns]

# Convert dates to ISO-compatible values.
date_columns = [
    "current_lease_start_date",
    "current_lease_end_date",
    "notice_due_date",
    "escalation_due",
    "initial_lease_commencement_date",
]
for col in date_columns:
    df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d")

# Convert numeric columns safely.
numeric_columns = [
    "basic_rent",
    "deposit_rent",
    "deposit_utilities",
    "key_deposit",
    "next_escalation_rental",
    "payment_score_overall",
    "payment_score_basic_rental",
    "payment_score_utilities",
    "tenure_months",
    "arrears_incidents_calc",
    "lease_longevity_months",
    "deposit_coverage_score",
    "tenant_risk_score",
]
for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors="coerce")

conn = sqlite3.connect(db_path)

with open(sql_path, "r", encoding="utf-8") as file:
    conn.executescript(file.read())

df.to_sql("tenant_risk_scorecard", conn, if_exists="append", index=False)

conn.commit()
conn.close()

print("Tenant risk scorecard imported successfully.")
print(f"Rows imported: {len(df)}")
print(f"Database: {db_path}")
