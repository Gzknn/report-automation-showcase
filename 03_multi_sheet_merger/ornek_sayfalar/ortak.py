"""Demo sayfa modüllerinin ortak şablonu (sentetik)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

SAYFALAR = ["Kanal Özeti", "Ürün Özeti"]


def sablon():
    wb = Workbook()
    wb.remove(wb.active)
    for ad in SAYFALAR:
        ws = wb.create_sheet(ad)
        ws["A1"] = ad
        ws["A1"].font = Font(bold=True, size=13)
        ws["A3"], ws["B3"] = "Kırılım", "Adet"
        for h in (ws["A3"], ws["B3"]):
            h.font = Font(bold=True, color="FFFFFF")
            h.fill = PatternFill("solid", fgColor="1F4E78")
        ws.column_dimensions["A"].width = 20
    return wb
