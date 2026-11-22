"""Pivot üreteci için SENTETİK test verisi üretir (gerçek veri içermez).

Kolon adları ve değerler tamamen uydurmadır; amaç üretecin boyut / ölçü
tespitini ve PivotTable çıktısını internetsiz denemektir.
"""
from pathlib import Path

import numpy as np
import pandas as pd

rng = np.random.default_rng(42)
N = 5_000
df = pd.DataFrame({
    "BASVURU_ID": np.arange(1, N + 1),                       # kimlik -> hariç tutulmalı
    "DONEM": rng.choice([202605, 202606, 202607], N),
    "KANAL": rng.choice(["Şube", "Mobil", "İnternet", "Çağrı Merkezi"], N, p=[.3, .4, .2, .1]),
    "URUN": rng.choice(["Ürün A", "Ürün B", "Ürün C"], N),
    "KARAR": rng.choice(["Onay", "Ret", "İnceleme"], N, p=[.55, .35, .10]),
    "SKOR": rng.integers(300, 1900, N),                      # ölçü
    "TUTAR": rng.gamma(2.0, 15_000, N).round(2),             # ölçü
})
cikti = Path(__file__).with_name("test_verisi.xlsx")
df.to_excel(cikti, index=False)
print(f"{len(df)} satırlık sentetik veri yazıldı: {cikti}")
