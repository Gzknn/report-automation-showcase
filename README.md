# Report Automation Showcase

Excel report automation tools I built in Python during my internship in a bank's
**Decision Support Systems** team. Monthly reports that used to be prepared by hand
(copy-paste from SQL, pivot, format) are generated end-to-end from the database
into a ready-to-send Excel file.

> **Note:** This is a cleaned, public version. All data, connection details,
> server addresses, schema/table names and business rules from the original
> project have been removed. Every example in this repo runs on **synthetic data**.
> Code comments are in Turkish (the language of the original team).

## What's inside

| Folder | What it does | Key ideas |
|---|---|---|
| [`01_template_report_engine`](01_template_report_engine) | Fills an **empty Excel template** from data using a readable rule dictionary | Tables are discovered from their headers (no hard-coded row numbers), rolling period window (latest month backwards), single counting point with a weight column, rules decoupled from code |
| [`02_excel_pivot_generator`](02_excel_pivot_generator) | Turns an Oracle query into a **real, drag-and-drop Excel PivotTable** | Auto-detects dimensions vs. measures from the data, aggregates to a summary cube in SQL (millions of rows → thousands), writes PivotTable OOXML parts directly into the `.xlsx` and validates them so Excel never opens a corrupt file |
| [`03_multi_sheet_merger`](03_multi_sheet_merger) | Runs each report page's own notebook/module and merges the results into **one workbook with formatting** | Loads `.ipynb` files as isolated modules, copies values + styles + merged cells + column widths, one failing page doesn't stop the report |

## Quick start

```bash
pip install -r requirements.txt

# 1) Template-driven report engine
cd 01_template_report_engine && python demo.py        # -> ornek_rapor.xlsx

# 2) Real Excel PivotTable from a query
cd ../02_excel_pivot_generator && python ornek_veri_uret.py && python pivot_table_generator.py   # -> pivot_table.xlsx

# 3) Merge several report pages into one workbook
cd ../03_multi_sheet_merger && python demo.py         # -> birlesik_rapor.xlsx
```

## Using a real database

Data access is isolated behind one boundary (`veri_kaynagi.kaynak_oku` / `VERI_KAYNAGI`).
Credentials are never written in code; they are read from environment variables:

```bash
export ORACLE_USER=...  ORACLE_PASSWORD=...  ORACLE_DSN="host:port/service"
pip install oracledb
```

There is **no silent fallback**: if the database fails, the error is raised instead of
quietly switching to an Excel file, so it is always clear which data a report came from.

## Tech

Python · pandas · openpyxl · Oracle (python-oracledb) · SQL · OOXML · Jupyter
