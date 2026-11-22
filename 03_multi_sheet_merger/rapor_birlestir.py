"""Birden çok rapor sayfasını, her birini KENDİ kodu/notebook'u ile üretip
TEK bir Excel dosyasında biçimiyle birlikte birleştirir.

Neden böyle?
- Her sayfanın notebook'u kendi başına çalışır ve test edilir. Birleşik
  rapor rapor mantığını tekrar yazmaz; bir sayfadaki düzeltme birleşik
  çıktıya kendiliğinden yansır.
- Notebook'larda aynı adlı ama içi farklı yardımcılar olabilir; bu yüzden
  her biri AYRI bir modül isim alanında yüklenir, çakışmazlar.
- Bir sayfa hata verirse rapor durmaz: o sayfa şablondaki boş hâliyle kalır,
  hata yazılır, diğer sayfalar ve dosya yine üretilir.
"""
import json
import shutil
import sys
import tempfile
import time
import traceback
import types
from copy import copy
from pathlib import Path

from openpyxl import load_workbook


# ----------------------------------------------------------------------------
# Modül yükleme: .py ya da .ipynb
# ----------------------------------------------------------------------------
def notebook_modulu_yukle(yol: Path, ad: str):
    """Bir .ipynb'nin KOD hücrelerini tek bir modül isim alanında çalıştırır.

    '%' / '!' ile başlayan IPython satırları atlanır; böylece notebook
    Jupyter'e gerek kalmadan sıradan bir modül gibi kullanılabilir.
    """
    nb = json.loads(Path(yol).read_text(encoding="utf-8"))
    parcalar = []
    for hucre in nb.get("cells", []):
        if hucre.get("cell_type") != "code":
            continue
        kaynak = hucre.get("source", "")
        if isinstance(kaynak, list):
            kaynak = "".join(kaynak)
        parcalar.append("\n".join(s for s in kaynak.splitlines()
                                  if not s.lstrip().startswith(("%", "!"))))
    modul = types.ModuleType(ad)
    modul.__file__ = str(yol)
    sys.modules[ad] = modul
    try:
        exec(compile("\n\n".join(parcalar), str(yol), "exec"), modul.__dict__)
    except Exception:
        sys.modules.pop(ad, None)
        raise
    return modul


def modul_yukle(yol: Path, ad: str):
    yol = Path(yol)
    if yol.suffix.lower() == ".ipynb":
        return notebook_modulu_yukle(yol, ad)
    import importlib.util
    spec = importlib.util.spec_from_file_location(ad, yol)
    modul = importlib.util.module_from_spec(spec)
    sys.modules[ad] = modul
    spec.loader.exec_module(modul)
    return modul


# ----------------------------------------------------------------------------
# Sayfa kopyalama (biçimiyle)
# ----------------------------------------------------------------------------
def sayfayi_kopyala(kaynak, hedef) -> None:
    """Doldurulmuş sayfayı (kaynak) birleşik dosyadaki sayfaya (hedef) taşır.

    openpyxl bir sayfayı başka bir dosyaya taşıyamaz; bu yüzden hücre hücre
    kopyalanır. Biçim nesneleri KOPYALANARAK verilir: her dosyanın kendi
    biçim tablosu vardır, kaynağın iç biçim numarası hedefte başka bir
    biçimi gösterirdi.
    """
    for aralik in list(hedef.merged_cells.ranges):
        hedef.unmerge_cells(str(aralik))

    for r in range(1, max(kaynak.max_row, hedef.max_row) + 1):
        for c in range(1, max(kaynak.max_column, hedef.max_column) + 1):
            k = kaynak.cell(row=r, column=c)
            h = hedef.cell(row=r, column=c)
            h.value = k.value
            h.font = copy(k.font)
            h.fill = copy(k.fill)
            h.border = copy(k.border)
            h.alignment = copy(k.alignment)
            h.protection = copy(k.protection)
            h.number_format = k.number_format

    for aralik in kaynak.merged_cells.ranges:
        hedef.merge_cells(str(aralik))
    for no, boyut in kaynak.row_dimensions.items():
        hedef.row_dimensions[no].height = boyut.height
        hedef.row_dimensions[no].hidden = boyut.hidden
    for harf, boyut in kaynak.column_dimensions.items():
        hedef.column_dimensions[harf].width = boyut.width
        hedef.column_dimensions[harf].hidden = boyut.hidden
    hedef.sheet_view.showGridLines = kaynak.sheet_view.showGridLines
    hedef.freeze_panes = kaynak.freeze_panes


# ----------------------------------------------------------------------------
# Birleşik rapor
# ----------------------------------------------------------------------------
def raporlari_birlestir(sayfa_modulleri: dict, sablon_yukle, cikti: Path) -> Path:
    """sayfa_modulleri: {sayfa adı: .py/.ipynb yolu}. Her modülde
    `generate_report(output_path)` fonksiyonu, o sayfayı dolduran bir Excel
    yazar. sablon_yukle(): bütün sayfaları boş hâliyle içeren çalışma kitabı.
    """
    basladi = time.time()
    gecici = Path(tempfile.mkdtemp(prefix="birlesik_"))
    uretilen, durumlar = {}, {}
    try:
        for i, (sayfa, yol) in enumerate(sayfa_modulleri.items()):
            print(f"\n### {sayfa} — kendi modülü çalıştırılıyor ({Path(yol).name})")
            dosya = gecici / f"sayfa_{i}.xlsx"
            try:
                modul = modul_yukle(yol, f"_sayfa_{i}")
                modul.generate_report(output_path=dosya)
                uretilen[sayfa] = dosya
                durumlar[sayfa] = "dolu"
            except Exception as exc:  # noqa: BLE001
                print(f"[UYARI] {sayfa} ÜRETİLEMEDİ ({type(exc).__name__}: {exc})")
                for satir in traceback.format_exc().strip().splitlines():
                    print(f"        {satir}")
                durumlar[sayfa] = f"BOŞ — {type(exc).__name__}"

        wb = sablon_yukle()
        for sayfa, dosya in uretilen.items():
            kaynak_wb = load_workbook(dosya)
            if sayfa not in kaynak_wb.sheetnames or sayfa not in wb.sheetnames:
                print(f"[UYARI] '{sayfa}' sayfası dosyada yok — birleştirilemedi.")
                durumlar[sayfa] += " (birleştirilemedi)"
                continue
            sayfayi_kopyala(kaynak_wb[sayfa], wb[sayfa])
        cikti = Path(cikti)
        wb.save(cikti)
    finally:
        shutil.rmtree(gecici, ignore_errors=True)

    print(f"\nBİRLEŞİK RAPOR: {cikti.resolve()}  ({time.time() - basladi:.1f} sn)")
    for sayfa, durum in durumlar.items():
        print(f"  - {sayfa:<14}: {durum}")
    return cikti
