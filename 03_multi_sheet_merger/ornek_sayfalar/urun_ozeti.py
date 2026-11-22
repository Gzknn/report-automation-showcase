"""Demo sayfa: ürün kırılımı (sentetik veri)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from ortak import sablon  # noqa: E402

VERI = {"Ürün A": 1510, "Ürün B": 930, "Ürün C": 510}


def generate_report(output_path):
    wb = sablon()
    ws = wb["Ürün Özeti"]
    for i, (k, v) in enumerate(VERI.items(), start=4):
        ws.cell(i, 1, k)
        ws.cell(i, 2, v).number_format = "#,##0"
    wb.save(output_path)
