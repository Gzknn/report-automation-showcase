"""Rapor motorunun uçtan uca demosu — tamamen SENTETİK veriyle.

1. Boş şablonu üretir (sablon_uret.py)
2. Sentetik başvuru verisi üretir
3. Kural sözlüğüyle (RAPOR) şablonu doldurur -> ornek_rapor.xlsx
"""
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

import rapor_motoru as m
from sablon_uret import sablon_uret

BURASI = Path(__file__).parent
SABLON = sablon_uret(BURASI / "ornek_sablon.xlsx")

# --- Sentetik veri: ön-gruplanmış bir küp (her satır bir kırılım + ADET) ---
rng = np.random.default_rng(7)
satirlar = []
for donem in [202605, 202606, 202607]:
    for kanal in ["Mobil", "İnternet", "Şube", "Çağrı Merkezi"]:
        for karar in ["Onay", "Ret", "İnceleme"]:
            satirlar.append({"DONEM": donem, "KANAL": kanal, "KARAR": karar,
                             "ADET": int(rng.integers(50, 900))})
VERI = pd.DataFrame(satirlar)

# --- Kural sözlüğü: hangi tablo hangi kaynaktan, hangi kırılımla dolar ---
def sabit_seviye(kolon, degerler):
    return {"satir": kolon, "mod": "sabit", "payda": "grand_total",
            "sabit": [{"etiket": d, "degerler": [d]} for d in degerler]}

RAPOR = {
    "surum": 4,
    "ad": "Demo Aylık Rapor",
    "kaynaklar": {"demo": {"ad": "Demo veri", "donemKolonu": "DONEM", "adetKolonu": "ADET"}},
    "gruplar": [],
    "sayfalar": {"Özet": {"paneller": {
        "Karar Dağılımı": {"ad": "Karar Dağılımı", "kaynak": "demo", "filtre": [],
                           "grandTotal": "gosterilen", "yerlesim": "sablon",
                           "agac": {"seviyeler": [sabit_seviye("KARAR", ["Onay", "Ret", "İnceleme"])]}},
        "Kanal Dağılımı": {"ad": "Kanal Dağılımı", "kaynak": "demo", "filtre": [],
                           "grandTotal": "gosterilen", "yerlesim": "sablon",
                           "agac": {"seviyeler": [sabit_seviye("KANAL", ["Mobil", "İnternet", "Şube", "Çağrı Merkezi"])]}},
    }}},
}

if __name__ == "__main__":
    m.rapor_uret(RAPOR, ["Özet"],
                 kaynak_oku=lambda kid, aylar=None: VERI,
                 sablon_yukle=lambda: load_workbook(SABLON),
                 cikti=BURASI / "ornek_rapor.xlsx")
