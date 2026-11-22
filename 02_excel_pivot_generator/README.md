# 02 · Query → real Excel PivotTable

Produces an `.xlsx` with a genuine PivotTable (field list, drag & drop in Excel), not a static table.

1. **Field detection:** a small sample determines type; distinct counts decide each
   column's role — *dimension* (few values), *excluded* (IDs, constants), or *measure* (`OLCU_ZORLA`).
2. **Summary cube:** grouping is done in SQL, so only the aggregated cube is downloaded,
   never millions of raw rows or row IDs.
3. **PivotTable injection:** openpyxl cannot write pivots, so `pivotCacheDefinition`,
   `pivotCacheRecords`, `pivotTable` XML and their relationships are written into the zip directly.
4. **Validation before saving:** index consistency between cache, shared items and pivot
   fields is checked; the target file is written only if everything passes.

```bash
python ornek_veri_uret.py        # synthetic test_verisi.xlsx
python pivot_table_generator.py  # -> pivot_table.xlsx
```
Set `VERI_KAYNAGI = "oracle"` and `ORACLE_SORGU` to run against a database.
