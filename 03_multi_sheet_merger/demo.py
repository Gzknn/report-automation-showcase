"""İki sayfayı kendi modülleriyle üretip tek dosyada birleştirir (sentetik)."""
import sys
from pathlib import Path

BURASI = Path(__file__).parent
sys.path.insert(0, str(BURASI / "ornek_sayfalar"))

from ortak import sablon  # noqa: E402
from rapor_birlestir import raporlari_birlestir  # noqa: E402

raporlari_birlestir(
    {"Kanal Özeti": BURASI / "ornek_sayfalar" / "kanal_ozeti.py",
     "Ürün Özeti": BURASI / "ornek_sayfalar" / "urun_ozeti.py"},
    sablon_yukle=sablon,
    cikti=BURASI / "birlesik_rapor.xlsx",
)
