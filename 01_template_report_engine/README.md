# 01 · Template-driven report engine

Fills an empty Excel template from data. You describe **what** each table shows in a
rule dictionary (`RAPOR`); the engine finds **where** the table is in the template.

- **Panel discovery:** a table is a header cell + period columns (`YYYYMM`) + row labels
  below it. No row/column numbers are hard-coded, so a template change needs no code change.
- **Rolling window:** the latest month in the data is found automatically; the report
  shows as many months backwards as the template has period columns (handles year roll-over).
- **Single counting point:** counts go through one function, so a pre-aggregated
  `COUNT(*)` column and raw rows are handled the same way.
- **Robust:** a table without a rule is skipped with a message, the report is still produced.

| File | |
|---|---|
| `rapor_motoru.py` | The engine |
| `veri_kaynagi.py` | Data boundary: Oracle (env-var credentials, chunked fetch + cache) or Excel (with pickle cache) |
| `sablon_uret.py` | Builds a synthetic demo template |
| `demo.py` | End-to-end run on synthetic data → `ornek_rapor.xlsx` |
