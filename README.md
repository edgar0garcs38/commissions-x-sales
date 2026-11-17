# Sales Commission Calculator & Payment Reconciliation

A data pipeline that calculates sales-agent commissions for an insurance brokerage and reconciles them against insurer payments. It replaced a manual, spreadsheet-based process.

## What it does

**Commission calculation** (`First Programs/main.py`)
- Unifies yearly portfolio exports into a single base (`DbUnification.py`)
- Identifies each agent's commission level (Level 1/2/3, special and collaborator tables)
- Classifies every transaction as issuance, renewal or plan change
- Finds the right commission rate per insurer, product and plan, using **fuzzy matching** (RapidFuzz) with strict token rules to avoid false matches between similar names
- Assigns the commission percentage and calculates the payment per policy

**Payment reconciliation** (`Version2-*` → `Version4.py`, the latest)
- Parses monetary values in mixed formats (`1.234,56`, `(123)`, `$`)
- Expands each policy into its installments (`cuotas`, e.g. `3/12`) and checks which are paid
- Detects **missing installments**
- Excludes canceled policies
- Classifies policies as fully paid, partially paid or unpaid
- Outputs an Excel workbook with `Cuota Status`, `Missing Cuotas`, `Payments Global` and `Match Results`

The numbered versions show how the logic evolved as new business rules came in.

## Tech stack
Python · pandas · NumPy · RapidFuzz · regex · Unicode normalization

## Run it
```bash
pip install -r requirements.txt
python Version4.py
```

> **Note:** The input files contain confidential client and company data, so they are not included in this repository. Only the source code is published.
