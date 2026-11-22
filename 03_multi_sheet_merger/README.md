# 03 · Multi-sheet report merger

Each report page lives in its own notebook and is tested on its own. The merger
runs every page's `generate_report()` in an **isolated module namespace**
(`.py` or `.ipynb`), then copies the filled sheets into one workbook — values, number
formats, fonts, fills, borders, merged cells, row heights and column widths.

If a page fails, it stays empty in the output, the full error is printed and the
remaining pages are still produced.

```bash
python demo.py   # -> birlesik_rapor.xlsx
```
