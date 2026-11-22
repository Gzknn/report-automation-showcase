"""Demo için boş bir Excel ŞABLONU üretir (sentetik, gerçek rapor değildir).

Motor tabloları başlıklarından bulduğu için şablonda kod tarafında tutulan
hiçbir satır numarası yoktur: başlık hücresi + yanında dönem sütunları
(YYYYMM) + altında satır etiketleri yeter.
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

AYLAR = [202605, 202606, 202607]
BASLIK = PatternFill("solid", fgColor="1F4E78")
BEYAZ = Font(bold=True, color="FFFFFF")
INCE = Side(style="thin", color="BFBFBF")
CERCEVE = Border(left=INCE, right=INCE, top=INCE, bottom=INCE)


def panel(ws, satir, sutun, baslik, etiketler):
    """Bir tablo: başlık (2 satır birleşik) + her ay için Adet/Oran sütunu."""
    ws.cell(satir, sutun, baslik)
    ws.merge_cells(start_row=satir, start_column=sutun, end_row=satir + 1, end_column=sutun)
    for i, ay in enumerate(AYLAR):
        c = sutun + 1 + 2 * i
        ws.cell(satir, c, ay)
        ws.merge_cells(start_row=satir, start_column=c, end_row=satir, end_column=c + 1)
        ws.cell(satir + 1, c, "Adet")
        ws.cell(satir + 1, c + 1, "Oran")
    son = sutun + 2 * len(AYLAR)
    for c in range(sutun, son + 1):
        for r in (satir, satir + 1):
            h = ws.cell(r, c)
            h.fill, h.font, h.border = BASLIK, BEYAZ, CERCEVE
            h.alignment = Alignment(horizontal="center", vertical="center")
    for j, etiket in enumerate(etiketler):
        r = satir + 2 + j
        ws.cell(r, sutun, etiket).border = CERCEVE
        if etiket == "Grand Total":
            ws.cell(r, sutun).font = Font(bold=True)
        for i in range(len(AYLAR)):
            adet = ws.cell(r, sutun + 1 + 2 * i)
            oran = ws.cell(r, sutun + 2 + 2 * i)
            adet.number_format, oran.number_format = "#,##0", "0.0%"
            adet.border = oran.border = CERCEVE
    ws.column_dimensions["A"].width = 22


def sablon_uret(yol: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Özet"
    ws["A1"] = "Aylık Başvuru Raporu (sentetik demo)"
    ws["A1"].font = Font(bold=True, size=14)
    panel(ws, 3, 1, "Karar Dağılımı", ["Onay", "Ret", "İnceleme", "Grand Total"])
    panel(ws, 10, 1, "Kanal Dağılımı", ["Mobil", "İnternet", "Şube", "Çağrı Merkezi", "Grand Total"])
    wb.save(yol)
    return yol


if __name__ == "__main__":
    print(sablon_uret(Path(__file__).with_name("ornek_sablon.xlsx")))
