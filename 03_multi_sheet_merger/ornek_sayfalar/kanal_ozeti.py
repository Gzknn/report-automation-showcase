"""Demo sayfa: kanal kırılımı (sentetik veri)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from ortak import sablon  # noqa: E402

VERI = {"Mobil": 1240, "İnternet": 860, "Şube": 640, "Çağrı Merkezi": 210}


def generate_report(output_path):
    wb = sablon()
    ws = wb["Kanal Özeti"]
    for i, (k, v) in enumerate(VERI.items(), start=4):
        ws.cell(i, 1, k)
        ws.cell(i, 2, v).number_format = "#,##0"
    wb.save(output_path)
