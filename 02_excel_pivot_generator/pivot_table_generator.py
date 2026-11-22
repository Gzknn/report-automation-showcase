# ============================================================================
# ORACLE -> GERÇEK EXCEL PİVOTTABLE
# ============================================================================
# Excel'de SÜRÜKLE-BIRAK yapabileceğiniz gerçek bir PivotTable üretir.
# Alan listesi kolon adlarından otomatik oluşur; satır, sütun, filtre ve
# değer yerleşimini Excel içinde istediğiniz gibi değiştirebilirsiniz.
#
# NEDEN DOSYADA BİR "Veri" SAYFASI VAR?
#   Excel'de bir PivotTable'ın çalışabilmesi için dosyada bir KAYNAK ARALIK
#   olmak zorundadır; alan listesi ve sürükle-bırak davranışı oradan gelir.
#   Ama bu aralığa HAM SATIRLAR yazılmaz: boyutlara göre gruplanmış bir
#   ÖZET KÜP yazılır (7 milyon satır -> genelde birkaç bin satır). Küpte
#   BASVURU_NO / MUSTERI_NO gibi satır kimlikleri bulunmaz.
#
# ALAN ADLARINI ELLE YAZMANIZ GEREKMEZ
#   Hangi kolonun kırılım (boyut), hangisinin sayı (ölçü) olduğu VERİDEN
#   okunur: küçük bir örnek çekilir, kolonların türüne ve kaç farklı değer
#   taşıdığına bakılır. Tespitin sonucu 3. bölümde tablo olarak yazılır;
#   katılmadığınız satır olursa ALAN_HARIC / BOYUT_ZORLA / OLCU_ZORLA ile
#   düzeltirsiniz.
#   İsterseniz VERI_SAYFASI_GIZLE = True ile bu sayfayı gizleyebilirsiniz;
#   pivot yine çalışır.
#
# BÖLÜMLER                                    sık değişir mi?
#   0  Kurulum ve ortam kontrolü              hayır (teşhis)
#   1  Ayarlar — sorgu ve tespit eşikleri     EVET, asıl çalışma alanı
#   2  Oracle bağlantısı                      hayır (ortam değişkenlerinden)
#   3  Alanların otomatik tespiti             hayır
#   4  Özet küp                               hayır
#   5  PivotTable XML üretimi                 hayır
#   6  Parçaları .xlsx içine yerleştirme      hayır
#   7  Rapor üretimi                          hayır
#
# Çalıştırma:  python pivot_table_generator.py
# ============================================================================


# ============================================================================
# 0. KURULUM VE ORTAM KONTROLÜ  (rapor için gerekli değil — teşhis amaçlı)
# ============================================================================
# Kurulum:  pip install -r requirements.txt   (Oracle için ayrıca: pip install oracledb)

# ============================================================================
# 1. AYARLAR
# ============================================================================
# Bu notebook, Excel'de SÜRÜKLE-BIRAK yapabileceğiniz GERÇEK bir PivotTable
# üretir. Alan listesi kolon adlarından otomatik oluşur; satır, sütun, filtre
# ve değer alanlarını Excel içinde istediğiniz gibi değiştirebilirsiniz.
import re
import shutil
import time
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from posixpath import dirname, join, normpath
from xml.sax.saxutils import escape

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

CIKTI_DOSYASI = Path("pivot_table.xlsx")

# --- KAYNAK SORGU --------------------------------------------------------
# Sorgunun etrafına parantez KOYMAYIN; kod gerektiğinde alt sorgu hâline
# getirir. WHERE ile daraltmak indirmeyi en çok azaltan adımdır.
ORACLE_SORGU = "SELECT * FROM <TABLO_ADI>   -- <<< DOLDURULACAK"

# --- ALANLAR: otomatik mı, elle mi? --------------------------------------
# True  -> BOYUTLAR ve OLCULER VERİDEN OKUNUR. Önce küçük bir örnek çekilir
#          (Oracle'da ROWNUM ile; tam veri indirilmez), her kolonun türüne ve
#          kaç farklı değer taşıdığına bakılıp rolü belirlenir. Ne seçtiğini
#          3. bölümde tablo hâlinde yazar.
# False -> aşağıdaki BOYUTLAR / OLCULER listeleri elle kullanılır.
OTOMATIK_ALANLAR = True

ORNEK_SATIR_SAYISI = 5000      # TÜR tespiti için çekilecek örnek satır sayısı

# Bir kolonda KAÇ FARKLI DEĞER olduğu nereden okunsun?
#   "kesin" -> tek bir COUNT(DISTINCT) sorgusuyla TÜM VERİDEN sayılır.
#              Bir tam tarama maliyeti vardır ama sonuç kesindir.
#   "ornek" -> sadece örnek satırlara bakılır. Hızlıdır ama YANILTABİLİR.
#
# NEDEN ÖRNEK YANILTIR? Oracle'da ROWNUM ile alınan örnek RASTGELE DEĞİLDİR,
# tablonun FİZİKSEL BAŞINDAN gelir. 8 milyonluk bir tablo tarihe göre
# sıralıysa ilk 5.000 satırın hepsi aynı aya düşer; COHORT "tek değer
# taşıyor" sanılıp elenir ve pivot ay kırılımını tamamen kaybeder.
# Bu yüzden varsayılan "kesin"dir.
KARDINALITE_MODU = "kesin"

# Bir kolon boyut sayılsın diye en fazla kaç farklı değer taşıyabilir?
# Bunu aşan SAYISAL kolonlar ölçü, diğerleri hariç olur.
MAKS_BOYUT_FARKLI_DEGER = 200

# Örnekteki satırların bu oranından fazlasında farklı değer taşıyan kolon
# KİMLİK sayılır (BASVURU_NO, MUSTERI_NO) ve tamamen dışarıda bırakılır.
KIMLIK_ORANI = 0.5

# Küp, boyutların farklı değer sayılarının ÇARPIMI kadar büyüyebilir. Bu
# sınır o patlamayı engeller; sınır dolarsa en az farklı değer taşıyanlar
# tutulur ve hangilerinin elendiği yazılır.
MAKS_BOYUT_SAYISI = 8

# KOD KOLONLARI: sayısal görünen ama toplanması ANLAMSIZ olan kolonlar.
# KARAR_KODU, SEHIR_KODU, TIP_KODU gibi alanlar Oracle'dan sayı
# olarak gelir; kural konmazsa "karar kodlarının toplamı" gibi bir ölçü
# üretilir. Adında bu parçalardan biri geçen sayısal kolon ölçü YAPILMAZ:
# az değerliyse boyut olur, çok değerliyse dışarıda bırakılır.
KOD_ISARETLERI = ("KOD", "CODE", "_NO", "TIP", "TYPE", "FLAG", "BANT", "_ID")

# Otomatik tespitin kararını değiştirmek isterseniz kolon adını buralara yazın.
ALAN_HARIC  = ()   # hiç kullanılmasın
BOYUT_ZORLA = ()   # her hâlükârda boyut olsun
OLCU_ZORLA  = ("TUTAR",)   # kolonun TOPLAMI da ölçü olarak eklensin (örnek veride TUTAR)

# --- SATIR SAYACI -------------------------------------------------------
# Küp ÖNCEDEN GRUPLANMIŞTIR: bir kırılımda kaç kaynak satır olduğu, ancak
# bir kolonda tutulursa bilinir. Bu kolon olmazsa Excel'de bir alanı
# Değerler'e sürüklediğinizde KÜP SATIRLARINI sayar, kaynak satırlarını
# değil — sayı sessizce yanlış çıkar.
#
# Bu yüzden sayaç, veriden okunmayan TEK alandır. Kapatabilirsiniz, ama o
# zaman pivotta satır sayısı göremezsiniz.
SAYAC_ADI = "SATIR_SAYISI"      # veride aynı adlı kolon varsa _2 eklenir

# Pivotun "Değerler" alanında SADECE bu sayaç bulunur. Kolonların toplamı
# otomatik ÖLÇÜ yapılmaz: alan listesinde sizin kendi kolonlarınız durur,
# pivot da her kırılımda kaç satır olduğunu gösterir.
#
# Bir kolonun TOPLAMINI da istiyorsanız adını OLCU_ZORLA'ya yazın; o zaman
# kolon kendi adıyla ölçü olur (Excel "Toplam TUTAR" diye gösterir).

# --- BOYUTLAR / OLCULER: sadece OTOMATIK_ALANLAR = False iken kullanılır --
# BOYUTLAR: bu kolonlara göre GROUP BY yapılır; alan listesinde bunlar görünür.
# DİKKAT: her satırda farklı olan bir kolon (BASVURU_NO gibi) KOYMAYIN —
# özet küp ham veri kadar büyür ve dosya şişer.
BOYUTLAR = ["AY", "KANAL", "NIHAI_KARAR", "SEGMENT"]

# OLCULER: pivotun "Değerler" alanına konabilecek sayılar.
#   ad    : Excel'de görünecek kolon adı
#   kolon : üzerinde işlem yapılacak kolon ("*" -> satır sayısı)
#   islem : "count" | "sum"
#
# NEDEN SADECE count VE sum? PivotTable, özet küpteki değerleri yeniden
# TOPLAYARAK çalışır. Toplam ve adet yeniden toplanabilir; ORTALAMA
# toplanamaz (ortalamaların ortalaması yanlış sonuç verir). Ortalama için
# SUM ve COUNT'u ayrı ölçü olarak ekleyip Excel'de "Hesaplanmış Alan" ile
# SUM/COUNT yazın — sonuç her kırılımda doğru olur.
OLCULER = [
    {"ad": "SATIR_SAYISI", "kolon": "*", "islem": "count"},
]

# Kaynak AYNI ADLI iki kolon döndürürse ne yapılsın? (sadece Excel yolu)
#   "yeniden_adlandir" -> ikincisi AD_2 olur, ne yapıldığı yazılır
#   "dur"              -> hata verilir
#
# ORACLE YOLUNDA bu ayar dikkate ALINMAZ, her zaman hata verilir: kolonu
# pandas'ta yeniden adlandırmak veri tabanındaki adı değiştirmez, küp
# sorgusu sonra var olmayan bir kolon arar. Sorguyu AS ile düzeltmek gerekir.
#
# Bu durum birleştirilmiş tablolarda (SELECT * FROM a JOIN b) çok olağandır:
# ortak anahtar iki kez gelir. Aynı adlı iki kolon varken df["AD"] tek bir
# kolon değil İKİ kolon döndürür ve kod "cannot convert the series to class
# int" gibi anlaşılmaz hatalarla kırılır.
TEKRARLI_KOLON = "yeniden_adlandir"

# Küpün kaynaktaki KAÇ SATIRI özetlediği yazılsın mı? ADET (COUNT(*))
# ölçüsünün toplamı, gruplamaya giren satır sayısına eşittir; bu sayı
# beklediğiniz satır sayısını tutmuyorsa sorgu istediğinizi okumuyordur.
KAPSAMA_RAPORU = True

# --- SAYFALAR ------------------------------------------------------------
VERI_SAYFASI  = "Veri"      # özet küp (pivotun kaynağı) — HAM VERİ DEĞİL
PIVOT_SAYFASI = "Pivot"     # PivotTable'ın durduğu sayfa
VERI_SAYFASI_GIZLE = False  # True -> kaynak sayfa gizlenir (pivot yine çalışır)

# --- PİVOTUN AÇILIŞ GÖRÜNÜMÜ --------------------------------------------
# Excel açıldığında hazır gelen yerleşim. Sürükleyerek istediğiniz gibi
# değiştirebilirsiniz; bu sadece başlangıç noktasıdır.
#   "otomatik" -> tespit edilen ilk boyut satıra, ikincisi sütuna, ilk ölçü
#                 değere konur (elle alan adı yazmanız gerekmez)
#   "KANAL"    -> alan adını kendiniz de yazabilirsiniz
#   None       -> o alan BOŞ başlar
PIVOT_SATIR  = "otomatik"
PIVOT_SUTUN  = "otomatik"
PIVOT_DEGER  = "otomatik"

# --- VERİ KAYNAĞI --------------------------------------------------------
# "oracle" -> Oracle'dan çekilir
# "excel"  -> EXCEL_TEST_DOSYASI okunur (yerleşimi denemek için)
VERI_KAYNAGI = "excel"   # örnek veriyle denemek için; gerçek kullanımda "oracle"
EXCEL_TEST_DOSYASI = Path("test_verisi.xlsx")

# Kolon adı SQL'de ayrılmış bir kelimeyse (LIMIT, ORDER, DATE...) True yapın.
# DİKKAT: Oracle'da tırnaklı ad BÜYÜK/küçük harfe duyarlıdır.
SQL_TIRNAK = False

# ============================================================================
# 2. ORACLE BAĞLANTISI
# ============================================================================
# Bağlantı bilgileri KODA YAZILMAZ; ortam değişkenlerinden okunur:
#   ORACLE_USER, ORACLE_PASSWORD, ORACLE_DSN   (ör. "sunucu:port/servis")
# Böylece dosya paylaşılsa ya da git'e gönderilse de şifre ve sunucu adresi
# dışarı çıkmaz. Excel modunda (VERI_KAYNAGI = "excel") hiçbiri gerekmez.
#
# KURULUM:  pip install oracledb
import os

try:
    import oracledb
except ImportError:  # Excel modunda oracledb olmadan da çalışsın
    oracledb = None


def db_connection():
    if oracledb is None:
        raise ConnectionError("oracledb kurulu değil: pip install oracledb")
    eksik = [a for a in ("ORACLE_USER", "ORACLE_PASSWORD", "ORACLE_DSN") if not os.environ.get(a)]
    if eksik:
        raise ConnectionError(f"Oracle bağlantı bilgisi eksik, ortam değişkenlerini tanımlayın: {eksik}")
    try:
        return oracledb.connect(user=os.environ["ORACLE_USER"],
                                password=os.environ["ORACLE_PASSWORD"],
                                dsn=os.environ["ORACLE_DSN"])
    except Exception as e:
        raise ConnectionError(f"Failed to establish database connection: {e}")


def fetch_data(query):
    conn = None
    try:
        conn = db_connection()
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [desc[0] for desc in cursor.description]
        data = cursor.fetchall()
        df = pd.DataFrame(data, columns=columns)
        cursor.close()
        return df
    except Exception as e:
        raise RuntimeError(f"Query execution failed: {e}")
    finally:
        if conn:
            conn.close()


# ----------------------------------------------------------------------------
# HIZ: PARÇALI ÇEKME + ÖNBELLEK
# ----------------------------------------------------------------------------
# fetch_data yukarıda AYNEN duruyor; aşağıdakiler onun yerine geçmez, yanına
# eklenir. İki darboğazı kapatırlar:
#
# 1) arraysize. python-oracledb varsayılan olarak sunucudan BİR SEFERDE 100
#    SATIR ister. 500.000 satırlık bir tablo için bu ~5.000 ağ gidiş-dönüşü
#    demektir; her gidiş-dönüş 20 ms sürse tek başına ~100 saniye eder ve
#    ekranda hiçbir şey görünmez. arraysize'ı büyütmek gidiş-dönüş sayısını
#    aynı oranda düşürür — tek satırlık, en büyük kazanç.
#
# 2) Aynı sorgunun İKİ KEZ çekilmesi. Birleşik raporlarda aynı tablo
#    birden çok sayfa için isteniyordu; önbellek bunu tek indirmeye düşürür.
#
# Ayrıca her parçada ilerleme yazılır: rapor "dondu mu, çalışıyor mu"
# sorusu ortadan kalkar.
ORACLE_ARRAYSIZE = 5000        # sunucudan bir seferde istenen satır sayısı
ORACLE_PARCA_BOYUTU = 100_000  # belleğe alınan parça büyüklüğü

# Aynı çalıştırma içinde tekrarlanan sorgular buraya yazılır.
_ORACLE_ONBELLEK: dict = {}


def fetch_data_hizli(query, onbellege_al: bool = True):
    """fetch_data ile AYNI sözleşme (DataFrame döndürür); sadece daha hızlı.

    Farkları: arraysize büyütülür, satırlar fetchmany ile parça parça
    alınır, ilerleme yazılır ve sonuç önbelleğe alınır.

    ÖNBELLEK NOTU: aynı sorgu ikinci kez istendiğinde KOPYA döndürülür.
    Böylece iki ayrı rapor aynı DataFrame'i paylaşıp birbirinin verisini
    değiştiremez; ilk çağıran ise kopya maliyeti ödemez.
    """
    anahtar = " ".join(str(query).split())
    if onbellege_al and anahtar in _ORACLE_ONBELLEK:
        print(f"[BİLGİ] Bu sorgu bu çalıştırmada ZATEN çekilmişti, önbellekten "
              f"kullanılıyor (tekrar indirilmedi).")
        return _ORACLE_ONBELLEK[anahtar].copy()

    conn = None
    try:
        conn = db_connection()
        cursor = conn.cursor()
        cursor.arraysize = ORACLE_ARRAYSIZE
        basladi = time.time()
        cursor.execute(query)
        columns = [desc[0] for desc in cursor.description]

        parcalar = []
        toplam = 0
        while True:
            rows = cursor.fetchmany(ORACLE_PARCA_BOYUTU)
            if not rows:
                break
            toplam += len(rows)
            parcalar.append(pd.DataFrame(rows, columns=columns))
            print(f"[BİLGİ]   ... {toplam:,} satır".replace(",", ".")
                  + f" ({time.time() - basladi:.0f} sn)")
        cursor.close()

        df = (pd.concat(parcalar, ignore_index=True) if parcalar
              else pd.DataFrame(columns=columns))
        print(f"[BİLGİ] {len(df):,} satır x {len(df.columns)} kolon".replace(",", ".")
              + f", {time.time() - basladi:.0f} saniyede çekildi "
              f"(arraysize={ORACLE_ARRAYSIZE}).")
        if onbellege_al:
            _ORACLE_ONBELLEK[anahtar] = df
        return df
    except Exception as e:
        raise RuntimeError(f"Query execution failed: {e}")
    finally:
        if conn:
            conn.close()

# ============================================================================
# 3. ALANLARIN OTOMATİK TESPİTİ
# ============================================================================
# BOYUTLAR ve OLCULER'i elle yazmak yerine VERİDEN okur.
#
# NASIL? Önce küçük bir ÖRNEK çekilir (Oracle'da ROWNUM <= N ile; 7 milyon
# satır İNDİRİLMEZ). Her kolonun türüne ve kaç farklı değer taşıdığına
# bakılarak rolü belirlenir:
#
#   kimlik   -> her satırda (neredeyse) farklı  : BASVURU_NO, MUSTERI_NO
#               HARİÇ TUTULUR. Boyut yapılsa küp ham veri kadar büyür,
#               ölçü yapılsa "başvuru numaralarının toplamı" gibi anlamsız
#               bir sayı çıkar.
#   ölçü     -> sayısal ve çok sayıda farklı değer : SKOR, TUTAR
#               SUM olarak eklenir.
#   boyut    -> az sayıda farklı değer : KANAL, NIHAI_KARAR, COHORT, *_FLAG
#               GROUP BY'a girer, Excel'de sürüklenebilir alan olur.
#   hariç    -> örnekte hep boş, ya da tek değer taşıyan kolonlar
#               (kırılım üretmezler, sadece alan listesini kalabalıklaştırır)
#
# ÖRNEKLEMENİN SINIRI: kardinalite ÖRNEKTEN tahmin edilir. Az sayıda farklı
# değer taşıyan kolonlar (KANAL, karar kodları, bayraklar) küçük örnekte de
# doğru görünür; bu ayrımın çalıştığı yer burasıdır. Ama tablonun tamamında
# örnekte görünenden daha fazla farklı değer olabilir — bu yüzden küp
# çekildikten SONRA gerçek büyüklük tekrar kontrol edilir (_kup_boyut_uyarisi).
#
# Tespitin sonucu bir TABLO olarak ekrana yazılır. Katılmadığınız bir satır
# olursa ALAN_HARIC / BOYUT_ZORLA / OLCU_ZORLA listelerine kolon adını
# yazmanız yeter; ya da OTOMATIK_ALANLAR = False yapıp eski elle yazma
# yöntemine dönersiniz.
import datetime
import decimal

ORNEK_SORGU_SABLONU = "SELECT * FROM (\n{sorgu}\n)\nWHERE ROWNUM <= {n}"

OTOMATIK = "otomatik"        # PIVOT_SATIR / PIVOT_SUTUN / PIVOT_DEGER sentinel'i
_ALAN_RAPORU: list = []      # son tespitin ayrıntısı (alan_raporu() ile yazdırılır)
_KAYNAK_SATIR: dict = {"sayi": None}   # kaynaktaki toplam satır (biliniyorsa)


def _ad(kolon: str) -> str:
    return f'"{kolon}"' if SQL_TIRNAK else str(kolon)


def _girintile(metin, bosluk="  "):
    return "\n".join(bosluk + s for s in str(metin).splitlines())


def _deger_turu(seri: pd.Series) -> str:
    """'sayi' | 'tarih' | 'metin'.

    dtype'a TEK BAŞINA GÜVENİLMEZ: oracledb sayıları int/float döndürür, ama
    araya bir tek None girdiğinde pandas kolonu 'object' yapar ve sayısal
    kolon metin gibi görünür. Bu yüzden tür, dolu gerçek değerlerden okunur.
    """
    if pd.api.types.is_numeric_dtype(seri):
        return "sayi"
    if pd.api.types.is_datetime64_any_dtype(seri):
        return "tarih"
    dolu = seri.dropna().head(200)
    if dolu.empty:
        return "metin"
    if all(isinstance(v, (int, float, decimal.Decimal)) and not isinstance(v, bool)
           for v in dolu):
        return "sayi"
    if all(isinstance(v, (datetime.datetime, datetime.date)) for v in dolu):
        return "tarih"
    return "metin"


def _tekrarli_kolonlari_coz(df: pd.DataFrame, nereden: str) -> pd.DataFrame:
    """Aynı adlı kolonları teke indirir.

    NEDEN GEREKLİ? pandas'ta iki kolonun adı aynıysa df["AD"] tek bir Series
    değil, İKİ KOLONLUK bir DataFrame döndürür. Sonra int(...) / .sum() gibi
    tek değer bekleyen her yer kırılır ve hata mesajı sebebi göstermez
    ("cannot convert the series to class int"). Sebebi burada, kaynağında
    söylenir.
    """
    adlar = [str(k) for k in df.columns]
    tekrar = sorted({a for a in adlar if adlar.count(a) > 1})
    if not tekrar:
        return df

    # ORACLE YOLUNDA YENİDEN ADLANDIRMA İŞE YARAMAZ. Kolonu pandas'ta
    # KANAL_2 yapmak veri tabanındaki adı değiştirmez; küp sorgusu sonra
    # "KANAL_2" diye var olmayan bir kolon arar. Üstelik aynı adlı iki
    # kolon içeren bir alt sorgudan kolon seçmek Oracle'da zaten
    # "ORA-00918: column ambiguously defined" verir. Tek çözüm sorguyu
    # düzeltmektir.
    if VERI_KAYNAGI.strip().lower() == "oracle":
        raise ValueError(
            f"{nereden}: AYNI ADLI kolon(lar) var: {tekrar}\n"
            "        ORACLE_SORGU bu kolonları iki kez döndürüyor; muhtemelen "
            "SELECT * ile birleştirilmiş\n"
            "        iki tabloda da aynı kolon var. Sorguda AS ile ayırın:\n"
            "            SELECT a.MUSTERI_NO AS MUSTERI_NO_A, "
            "b.MUSTERI_NO AS MUSTERI_NO_B, ...\n"
            "        ya da SELECT * yerine istediğiniz kolonları tek tek yazın.\n"
            "        (Yeniden adlandırma burada işe yaramaz: küp sorgusu "
            "kolonu veri tabanındaki\n"
            "        adıyla aramak zorundadır.)")

    if TEKRARLI_KOLON.strip().lower() == "dur":
        raise ValueError(
            f"{nereden}: AYNI ADLI kolon(lar) var: {tekrar}\n"
            "        Kolon adlarını ayırın, ya da "
            'TEKRARLI_KOLON = "yeniden_adlandir" yapın.')

    yeni, sayac = [], {}
    for a in adlar:
        if adlar.count(a) == 1:
            yeni.append(a)
            continue
        sayac[a] = sayac.get(a, 0) + 1
        yeni.append(a if sayac[a] == 1 else f"{a}_{sayac[a]}")
    df = df.copy()
    df.columns = yeni
    print(f"[UYARI] {nereden}: aynı adlı kolon(lar) vardı: {tekrar}")
    print(f"        Yeniden adlandırıldı: "
          f"{[y for a, y in zip(adlar, yeni) if a != y]}")
    print("        Sorguda AS ile ayırmak daha iyidir; hangi kolonun hangi "
          "tablodan geldiği belli olur.")
    return df


def _benzersiz_ad(taban: str, kullanilan) -> str:
    """Kullanılmayan bir ad üretir: ADET -> ADET_2 -> ADET_3 ..."""
    ad, k = taban, 2
    while ad in kullanilan:
        ad = f"{taban}_{k}"
        k += 1
    return ad


def _ornek_veri(n: int = None) -> pd.DataFrame:
    """Alan tespiti için küçük bir örnek okur (tam veriyi indirmez)."""
    n = int(n or ORNEK_SATIR_SAYISI)
    if VERI_KAYNAGI.strip().lower() == "oracle":
        sorgu = ORNEK_SORGU_SABLONU.format(
            sorgu=_girintile(ORACLE_SORGU.strip().rstrip(";")), n=n)
        print(f"[BİLGİ] Alanları tespit etmek için örnek çekiliyor "
              f"(en fazla {n:,} satır — tam veri İNDİRİLMİYOR).".replace(",", "."))
        return _tekrarli_kolonlari_coz(fetch_data_hizli(sorgu), "Örnek sorgu")

    if not EXCEL_TEST_DOSYASI.exists():
        raise FileNotFoundError(
            f"Excel test dosyası yok: {EXCEL_TEST_DOSYASI.resolve()}\n"
            f"        Çalışma dizini: {Path.cwd()}")
    print(f"[BİLGİ] Alanları tespit etmek için örnek okunuyor "
          f"(en fazla {n:,} satır).".replace(",", "."))
    return _tekrarli_kolonlari_coz(pd.read_excel(EXCEL_TEST_DOSYASI, nrows=n),
                                   "Excel test dosyası")


def _kardinalite_sorgusu(kolonlar) -> str:
    """Her kolonun farklı değer sayısını TEK SORGUDA, TÜM VERİDEN sayar."""
    parca = ["COUNT(*) AS S_"]
    for i, k in enumerate(kolonlar):
        parca.append(f"COUNT(DISTINCT {_ad(k)}) AS D{i}_")
        parca.append(f"COUNT({_ad(k)}) AS N{i}_")
    icerik = ORACLE_SORGU.strip().rstrip(";")
    return (f"SELECT {', '.join(parca)}\n"
            f"FROM (\n{_girintile(icerik)}\n)")


def _tek_deger(x, nerede: str):
    """Tek bir değer bekleyen yerlerde anlaşılır hata verir.

    Bir kolon adı iki kolona birden uyduğunda pandas Series/DataFrame
    döndürür; int(...) çağrısı "cannot convert the series to class int" gibi
    sebebi göstermeyen bir hata verir. Burada sebebi söylenir.
    """
    if isinstance(x, (pd.Series, pd.DataFrame)):
        raise ValueError(
            f"{nerede}: tek bir değer bekleniyordu ama {len(x)} değer geldi.\n"
            "        Sebebi genelde AYNI ADLI İKİ KOLON olmasıdır; sorguda "
            "kolon adlarını AS ile ayırın.")
    return x


def _kesin_kardinalite(kolonlar) -> tuple:
    """(satir_sayisi, {kolon: farkli}, {kolon: dolu}) — tam taramayla."""
    sorgu = _kardinalite_sorgusu(kolonlar)
    print(f"[BİLGİ] Kardinalite TÜM VERİDEN sayılıyor (tek sorgu, tek tarama).")
    print(f"        Yavaş gelirse KARDINALITE_MODU = \"ornek\" yapabilirsiniz — "
          "ama örnek yanıltabilir (aşağıdaki nota bakın).")
    basladi = time.time()
    df = fetch_data_hizli(sorgu)
    satir = int(_tek_deger(df["S_"].iloc[0], "satır sayısı"))
    farkli = {k: int(_tek_deger(df[f"D{i}_"].iloc[0], f"'{k}' farklı değer sayısı"))
              for i, k in enumerate(kolonlar)}
    dolu = {k: int(_tek_deger(df[f"N{i}_"].iloc[0], f"'{k}' dolu değer sayısı"))
            for i, k in enumerate(kolonlar)}
    print(f"[BİLGİ] Kaynakta {satir:,} satır var".replace(",", ".")
          + f" ({time.time() - basladi:.0f} saniyede sayıldı).")
    return satir, farkli, dolu


def _rol_ata(ad: str, tur: str, farkli: int, satir: int) -> tuple:
    """Bir kolonun rolünü ve gerekçesini döndürür: (rol, sebep)."""
    oran = farkli / satir if satir else 0.0

    # Elle verilen kararlar her şeyin önünde gelir.
    if ad in ALAN_HARIC:
        return "hariç", "ALAN_HARIC listesinde"
    if ad in BOYUT_ZORLA:
        return "boyut", "BOYUT_ZORLA listesinde"
    if ad in OLCU_ZORLA:
        return "ölçü", "OLCU_ZORLA listesinde"

    if farkli == 0:
        return "hariç", "hiç dolu değer yok"
    if farkli == 1:
        return "hariç", "tek değer taşıyor — kırılım üretmez"

    # Kimlik kontrolü TÜR kontrolünden ÖNCE gelmeli: BASVURU_NO sayısaldır,
    # sırf sayısal diye ölçü yapılırsa "başvuru numaraları toplamı" çıkar.
    if oran >= KIMLIK_ORANI:
        return "hariç", f"satırların %{oran * 100:.0f}'ında farklı — kimlik kolonu"
    if tur == "tarih" and farkli > MAKS_BOYUT_FARKLI_DEGER:
        return "hariç", f"tarih/saat, {farkli} farklı değer — küpü şişirir"

    # KOD kolonu mu? KARAR_KODU sayı olarak gelir ama TOPLANMAZ; sadece
    # "sayısal" olduğuna bakılırsa anlamsız bir ölçü üretilir.
    kod_gibi = any(p in ad.upper() for p in KOD_ISARETLERI)
    if tur == "sayi" and kod_gibi and farkli > MAKS_BOYUT_FARKLI_DEGER:
        return "hariç", (f"adı kod/bayrak gibi, {farkli} farklı değer — "
                         "toplanması anlamsız olur")
    if tur == "sayi" and not kod_gibi and farkli > MAKS_BOYUT_FARKLI_DEGER:
        # Otomatik ÖLÇÜ yapılmaz: pivotun Değerler alanında sadece satır
        # sayacı durur, alan listesinde de sizin kolonlarınız. Bu kolonun
        # toplamını istiyorsanız adını OLCU_ZORLA'ya yazın.
        return "hariç", (f"sayısal, {farkli} farklı değer — kırılım için fazla "
                         "(toplamını istiyorsanız OLCU_ZORLA)")
    if farkli > MAKS_BOYUT_FARKLI_DEGER:
        return "hariç", (f"{farkli} farklı değer > MAKS_BOYUT_FARKLI_DEGER "
                         f"({MAKS_BOYUT_FARKLI_DEGER})")
    return "boyut", f"{farkli} farklı değer"


def alanlari_kesfet(ornek: pd.DataFrame = None) -> tuple:
    """(boyutlar, olculer) üretir ve tespit tablosunu yazar.

    TÜR küçük bir örnekten okunur (tür satır sırasına göre değişmez).
    KAÇ FARKLI DEĞER olduğu ise KARDINALITE_MODU'na bağlıdır:
      "kesin" -> tek COUNT(DISTINCT) sorgusuyla tüm veriden (varsayılan)
      "ornek" -> sadece örnek satırlardan (hızlı ama yanıltabilir)
    """
    global _ALAN_RAPORU
    if ornek is None:
        ornek = _ornek_veri()
    ornek_satir = len(ornek)
    if ornek_satir == 0:
        raise RuntimeError(
            "Örnek BOŞ — sorgu hiç satır döndürmedi, alanlar tespit edilemez.\n"
            "        ORACLE_SORGU'yu ve WHERE koşullarını kontrol edin.")

    kolonlar = [str(k) for k in ornek.columns]
    kesin = (KARDINALITE_MODU.strip().lower() == "kesin"
             and VERI_KAYNAGI.strip().lower() == "oracle")
    if kesin:
        satir, farkli_harita, _ = _kesin_kardinalite(kolonlar)
    else:
        satir = ornek_satir
        farkli_harita = {k: int(_tek_deger(ornek[k].nunique(dropna=True),
                                           f"'{k}' farklı değer sayısı"))
                         for k in kolonlar}
        if VERI_KAYNAGI.strip().lower() == "excel":
            # Excel yolunda dosyanın tamamı zaten okunabiliyor; örnek
            # kısıtlaması sadece Oracle tarafında anlamlı.
            tam = pd.read_excel(EXCEL_TEST_DOSYASI)
            satir = len(tam)
            farkli_harita = {k: int(_tek_deger(tam[k].nunique(dropna=True),
                                               f"'{k}' farklı değer sayısı"))
                             for k in kolonlar}

    rapor = []
    for kolon in kolonlar:
        tur = _deger_turu(ornek[kolon])
        farkli = farkli_harita[kolon]
        rol, sebep = _rol_ata(kolon, tur, farkli, satir)
        if kesin and ornek[kolon].dropna().empty and farkli > 0:
            # Örnekte hep boş ama veride dolu: tür okunamadı, metin varsayılır.
            sebep += " (tür okunamadı: örnekte bu kolon boştu)"
        rapor.append({"alan": kolon, "rol": rol, "farkli": farkli,
                      "tur": tur, "sebep": sebep})

    # Boyut sayısı sınırı: küp, boyutların farklı değer sayılarının ÇARPIMI
    # kadar büyüyebilir. Sınırı aşarsak en az farklı değer taşıyanları
    # tutarız (küpü en az büyütenler onlardır), kalanları raporda gösteririz.
    boyut_satirlari = [r for r in rapor if r["rol"] == "boyut"]
    if len(boyut_satirlari) > MAKS_BOYUT_SAYISI:
        tutulan = {r["alan"] for r in
                   sorted(boyut_satirlari, key=lambda r: (r["farkli"], r["alan"]))
                   [:MAKS_BOYUT_SAYISI]}
        for r in boyut_satirlari:
            if r["alan"] not in tutulan:
                r["rol"] = "hariç"
                r["sebep"] = (f"MAKS_BOYUT_SAYISI = {MAKS_BOYUT_SAYISI} doldu "
                              "(en az farklı değer taşıyanlar tutuldu)")

    # Kolon sırası korunur: alan listesi tablodaki sırayla görünsün.
    boyutlar = [r["alan"] for r in rapor if r["rol"] == "boyut"]
    olcu_alanlari = [r["alan"] for r in rapor if r["rol"] == "ölçü"]

    # Değerler alanında SADECE satır sayacı durur. Küp önceden gruplandığı
    # için bir kırılımda kaç kaynak satır olduğu ancak bir kolonda tutulursa
    # bilinir; bu yüzden sayaç veriden okunmayan tek alandır.
    sayac = _benzersiz_ad(str(SAYAC_ADI), set(boyutlar))
    if sayac != str(SAYAC_ADI):
        print(f"[UYARI] Veride '{SAYAC_ADI}' adlı bir kolon var; satır sayacı "
              f"'{sayac}' olarak adlandırıldı.")
    olculer = [{"ad": sayac, "kolon": "*", "islem": "count"}]

    # Kolon toplamı SADECE açıkça istenirse eklenir; kendi adıyla eklenir,
    # "TUTAR_TOPLAM" gibi yeni bir ad uydurulmaz (Excel "Toplam TUTAR" der).
    for a in olcu_alanlari:
        olculer.append({"ad": a, "kolon": a, "islem": "sum"})

    _ALAN_RAPORU = rapor
    _KAYNAK_SATIR["sayi"] = satir if kesin or VERI_KAYNAGI.strip().lower() == "excel" else None
    alan_raporu()

    if not boyutlar:
        raise RuntimeError(
            "Hiçbir kolon BOYUT olarak seçilemedi — pivot kurulamaz.\n"
            "        Yukarıdaki tabloda sebepler yazıyor. Bir kolonu zorla "
            "boyut yapmak için BOYUT_ZORLA listesine adını yazın,\n"
            "        ya da MAKS_BOYUT_FARKLI_DEGER değerini yükseltin.")
    return boyutlar, olculer


def alan_raporu() -> None:
    """Son tespitin tablosunu yazar."""
    if not _ALAN_RAPORU:
        print("[BİLGİ] Henüz alan tespiti yapılmadı.")
        return

    genislik = max([len(r["alan"]) for r in _ALAN_RAPORU] + [5])
    n = _KAYNAK_SATIR.get("sayi")
    kaynak = (f"{n:,} satırın TAMAMINDAN sayıldı".replace(",", ".") if n
              else f"sadece {ORNEK_SATIR_SAYISI:,} satırlık ÖRNEKTEN tahmin edildi"
                   .replace(",", "."))
    print("\n" + "-" * 78)
    print(f"ALAN TESPİTİ  ({kaynak})")
    print("-" * 78)
    print(f"{'ALAN'.ljust(genislik)}  {'ROL':6}  {'FARKLI':>7}  {'TÜR':6}  SEBEP")
    for r in sorted(_ALAN_RAPORU, key=lambda r: ({"boyut": 0, "ölçü": 1, "hariç": 2}[r["rol"]],)):
        print(f"{r['alan'].ljust(genislik)}  {r['rol']:6}  {r['farkli']:>7}  "
              f"{r['tur']:6}  {r['sebep']}")
    print("-" * 78)

    boyut = [r for r in _ALAN_RAPORU if r["rol"] == "boyut"]
    tahmin = 1
    for r in boyut:
        tahmin *= max(r["farkli"], 1)
    print(f"Boyut: {len(boyut)}  ·  Ölçü: {sum(1 for r in _ALAN_RAPORU if r['rol'] == 'ölçü')}"
          f"  ·  Hariç: {sum(1 for r in _ALAN_RAPORU if r['rol'] == 'hariç')}")
    print(f"Küpün ÜST SINIRI (boyutların çarpımı): {tahmin:,} satır".replace(",", "."))
    print("  Gerçek küp bundan küçük olur (her kombinasyon veride bulunmaz).")
    if _KAYNAK_SATIR.get("sayi") is None:
        print("[UYARI] Farklı değer sayıları ÖRNEKTEN tahmin edildi. Oracle'da örnek")
        print("        tablonun fiziksel başından gelir, rastgele değildir: tablo")
        print("        tarihe göre sıralıysa ay kolonu 'tek değerli' sanılıp elenebilir.")
        print("        Kesin sonuç için: KARDINALITE_MODU = \"kesin\".")
    if tahmin > KUP_UYARI_ESIGI:
        print(f"[UYARI] Üst sınır {KUP_UYARI_ESIGI:,} eşiğini aşıyor.".replace(",", ".")
              + " MAKS_BOYUT_SAYISI'nı düşürmeyi veya")
        print("        gereksiz bir boyutu ALAN_HARIC'e eklemeyi düşünün.")
    # Sırf eşik yüzünden elenenler için somut eşik değeri öner: bu kolonlar
    # (karar kodları gibi) çoğu zaman gerçekten istenen kırılımlardır.
    esik_kurbani = [r for r in _ALAN_RAPORU
                    if r["rol"] == "hariç" and MAKS_BOYUT_FARKLI_DEGER < r["farkli"]
                    and "kimlik" not in r["sebep"]]
    if esik_kurbani:
        gerekli = max(r["farkli"] for r in esik_kurbani)
        print("Sırf farklı değer sayısı yüzünden elenenler: "
              + ", ".join(f"{r['alan']} ({r['farkli']})" for r in esik_kurbani))
        print(f"  Bunları BOYUT yapmak için: MAKS_BOYUT_FARKLI_DEGER = {gerekli} "
              "(ya da adlarını BOYUT_ZORLA'ya yazın).")
        print("  DİKKAT: küp büyür — üstteki üst sınır bu kolonların farklı değer "
              "sayısıyla ÇARPILIR.")
    print("Değiştirmek için: ALAN_HARIC / BOYUT_ZORLA / OLCU_ZORLA listeleri, "
          "ya da OTOMATIK_ALANLAR = False.")
    print("-" * 78 + "\n")


def alanlari_belirle() -> tuple:
    """BOYUTLAR ve OLCULER'i kesinleştirir (otomatik ya da elle)."""
    global BOYUTLAR, OLCULER
    if not OTOMATIK_ALANLAR:
        print("[BİLGİ] OTOMATIK_ALANLAR = False — elle yazılan BOYUTLAR/OLCULER "
              "kullanılıyor.")
        return BOYUTLAR, OLCULER
    BOYUTLAR, OLCULER = alanlari_kesfet()
    return BOYUTLAR, OLCULER


def _acilis_yerlesimi(boyutlar, olculer) -> tuple:
    """'otomatik' yazan açılış alanlarını gerçek alan adlarına çevirir.

    None yazılırsa o alan BOŞ başlar — Excel'de sürükleyerek doldurursunuz.
    """
    def coz(deger, secenekler, sira):
        if deger is None or str(deger).strip().lower() != OTOMATIK:
            return deger
        return secenekler[sira] if len(secenekler) > sira else None

    olcu_adlari = [o["ad"] for o in olculer]
    return (coz(PIVOT_SATIR, boyutlar, 0),
            coz(PIVOT_SUTUN, boyutlar, 1),
            coz(PIVOT_DEGER, olcu_adlari, 0))


# ============================================================================
# 4. ÖZET KÜP  —  pivotun kaynağı
# ============================================================================
# PivotTable'ın Excel içinde çalışabilmesi için dosyada bir KAYNAK ARALIK
# olması gerekir; Excel alan listesini ve sürükle-bırak davranışını bu
# aralıktan üretir. Ama bu aralığa HAM SATIRLAR yazılmaz.
#
# Bunun yerine, tespit edilen BOYUTLAR'a göre gruplanmış bir ÖZET KÜP yazılır:
#   ham:  7.000.000 satır, müşteri/başvuru kırılımında
#   küp:  BOYUTLAR'ın farklı değer kombinasyonu kadar satır (genelde binler)
#
# Küpte BASVURU_NO, MUSTERI_NO gibi satır kimlikleri YOKTUR. Pivot yine de
# tam serbestlikle çalışır, çünkü ölçüler yeniden toplanabilir niteliktedir.
ISLEM_SQL = {
    "count": lambda k: "COUNT(*)" if k == "*" else f"COUNT({k})",
    "sum":   lambda k: f"SUM({k})",
}

# Küp bu satır sayısını aşarsa uyarılır. Aşması, BOYUTLAR'a her satırda farklı
# olan bir kolon (BASVURU_NO, MUSTERI_NO, tarih-saat...) konduğunun işaretidir;
# o zaman küp ham veri kadar büyür, dosya şişer ve pivot yavaşlar.
KUP_UYARI_ESIGI = 200_000


# (_ad ve _girintile 3. bölümde tanımlanır — sorgu metnini ikisi de kullanır.)


def _kup_sorgusu() -> str:
    """BOYUTLAR'a göre gruplayan, OLCULER'i hesaplayan sorguyu üretir."""
    if not BOYUTLAR:
        raise ValueError("BOYUTLAR boş olamaz — en az bir kırılım kolonu gerekir.")
    if not OLCULER:
        raise ValueError("OLCULER boş olamaz — en az bir ölçü gerekir.")

    # Ölçü adı bir boyut adıyla ya da başka bir ölçüyle aynı olamaz: sorgu
    # AYNI ADLI İKİ KOLON döndürür, df["AD"] tek kolon yerine iki kolon verir
    # ve kod "cannot convert the series to class int" ile kırılır.
    adlar = [str(b) for b in BOYUTLAR] + [str(o["ad"]) for o in OLCULER]
    cakisan = sorted({a for a in adlar if adlar.count(a) > 1})
    if cakisan:
        raise ValueError(
            f"Şu ad(lar) hem boyut hem ölçü olarak ya da iki kez kullanılmış: "
            f"{cakisan}\n"
            "        Küpte aynı adlı iki kolon olamaz. OLCULER içindeki 'ad' "
            "alanını değiştirin (ör. 'ADET' -> 'SATIR_ADEDI').")

    olcu_parcalari = []
    for o in OLCULER:
        islem = o.get("islem", "count")
        if islem not in ISLEM_SQL:
            raise ValueError(
                f"{o.get('ad')!r}: '{islem}' desteklenmiyor. Geçerli: {sorted(ISLEM_SQL)}.\n"
                "        Ortalama için SUM ve COUNT'u ayrı ölçü olarak ekleyip "
                "Excel'de Hesaplanmış Alan kullanın (bkz. 1. bölümdeki not).")
        kolon = o.get("kolon", "*")
        ifade = ISLEM_SQL[islem](kolon if kolon == "*" else _ad(kolon))
        olcu_parcalari.append(f'{ifade} AS {_ad(o["ad"])}')

    secim = ", ".join(_ad(b) for b in BOYUTLAR)
    icerik = ORACLE_SORGU.strip().rstrip(";")
    return (f"SELECT {secim}, {', '.join(olcu_parcalari)}\n"
            f"FROM (\n{_girintile(icerik)}\n)\n"
            f"GROUP BY {secim}")


def kup_verisi() -> pd.DataFrame:
    """Özet küpü üretir. Oracle modunda gruplama veri tabanında yapılır."""
    if VERI_KAYNAGI.strip().lower() == "oracle":
        sorgu = _kup_sorgusu()
        print(f"[BİLGİ] Özet küp sorgusu:\n{_girintile(sorgu, '        ')}")
        df = _tekrarli_kolonlari_coz(fetch_data_hizli(sorgu), "Küp sorgusu")
        print(f"[BİLGİ] Küp: {len(df):,} satır x {len(df.columns)} kolon "
              "(HAM VERİ İNDİRİLMEDİ).".replace(",", "."))
        _kapsama_raporu(df)
        _kup_boyut_uyarisi(df)
        return df

    # Excel test yolu: gruplamayı pandas yapar, sonuç aynı biçimdedir.
    if not EXCEL_TEST_DOSYASI.exists():
        raise FileNotFoundError(
            f"Excel test dosyası yok: {EXCEL_TEST_DOSYASI.resolve()}\n"
            f"Çalışma dizini: {Path.cwd()}")
    ham = _tekrarli_kolonlari_coz(pd.read_excel(EXCEL_TEST_DOSYASI), "Excel test dosyası")
    _kolon_kontrol(ham, BOYUTLAR + [o["kolon"] for o in OLCULER if o.get("kolon", "*") != "*"])
    grup = ham.groupby(BOYUTLAR, dropna=False)
    parcalar = {}
    for o in OLCULER:
        kolon, islem = o.get("kolon", "*"), o.get("islem", "count")
        parcalar[o["ad"]] = (grup.size() if kolon == "*" and islem == "count"
                             else grup[kolon].agg(islem))
    df = pd.DataFrame(parcalar).reset_index()
    print(f"[BİLGİ] Küp (Excel yolu): {len(df)} satır x {len(df.columns)} kolon.")
    _kapsama_raporu(df)
    _kup_boyut_uyarisi(df)
    return df


def _kapsama_raporu(df: pd.DataFrame) -> None:
    """Küpün kaynaktaki KAÇ SATIRI özetlediğini yazar.

    "8 milyon satırı 5 saniyede okudu" şüphesinin cevabı burada. Kod
    8 milyon satırı İNDİRMEZ: gruplamayı Oracle yapar, ağdan sadece küp
    iner. O yüzden süre indirme süresi değil, Oracle'ın tarama süresidir.

    Ama bunu SÖYLEMEK yetmez, GÖSTERMEK gerekir: ADET = COUNT(*) ölçüsünün
    küp genelindeki toplamı, gruplamaya giren satır sayısına EŞİTTİR. Bu
    sayı beklediğiniz satır sayısını tutmuyorsa sorgu istediğinizi
    okumuyordur (ör. ORACLE_SORGU'da unutulmuş bir WHERE ya da ROWNUM).
    """
    if not KAPSAMA_RAPORU:
        return
    sayac = next((o["ad"] for o in OLCULER
                  if o.get("islem") == "count" and o.get("kolon", "*") == "*"), None)
    if sayac is None or sayac not in df.columns:
        print("[BİLGİ] Kapsama raporu atlandı: OLCULER içinde satır sayacı yok.")
        return

    kapsanan = int(_tek_deger(df[sayac].sum(), f"'{sayac}' toplamı"))
    print(f"[BİLGİ] KAPSAMA: küp {kapsanan:,} kaynak satırı özetliyor.".replace(",", "."))
    beklenen = _KAYNAK_SATIR.get("sayi")
    if beklenen is None:
        print("        Bu sayı beklediğiniz satır sayısını tutmuyorsa sorgu "
              "istediğinizi okumuyordur.")
    elif kapsanan == beklenen:
        print(f"        Kaynaktaki satır sayısıyla EŞLEŞİYOR ({beklenen:,}) — "
              "hiçbir satır dışarıda kalmadı.".replace(",", "."))
    else:
        print(f"[UYARI] Kaynakta {beklenen:,} satır sayılmıştı, küp {kapsanan:,} "
              "satır özetliyor.".replace(",", "."))
        print("        Aradaki fark, gruplamaya girmeyen satırlar demektir; "
              "ORACLE_SORGU'yu kontrol edin.")


def _kup_boyut_uyarisi(df: pd.DataFrame) -> None:
    """Küp beklenenden büyükse sebebini işaret et."""
    if len(df) <= KUP_UYARI_ESIGI:
        return
    print(f"[UYARI] Özet küp {len(df):,} satır — beklenenden büyük.".replace(",", "."))
    print("        Sebebi genelde BOYUTLAR içinde HER SATIRDA FARKLI olan bir kolon "
          "bulunmasıdır.")
    for b in BOYUTLAR:
        if b in df.columns:
            farkli = df[b].nunique()
            if farkli > len(df) * 0.5:
                print(f"        Şüpheli: '{b}' -> {farkli:,} farklı değer"
                      .replace(",", "."))
    print("        Bu kolonu BOYUTLAR'dan çıkarın; dosya küçülür, pivot hızlanır.")


def _kolon_kontrol(df: pd.DataFrame, kolonlar) -> None:
    eksik = [k for k in kolonlar if k not in df.columns]
    if not eksik:
        return
    import difflib
    oneri = {k: difflib.get_close_matches(k, [str(c) for c in df.columns], n=3, cutoff=0.6)
             for k in eksik}
    raise ValueError(f"Şu kolon(lar) veride YOK: {eksik}\n"
                     f"        En benzer adlar: {oneri}\n"
                     f"        Verideki kolonlar: {list(df.columns)}")

# ============================================================================
# 5. GERÇEK EXCEL PİVOTTABLE ÜRETİMİ
# ============================================================================
# openpyxl PivotTable'ı sıfırdan YAZAMAZ (sadece var olanı okuyabilir). Bu
# yüzden .xlsx (aslında bir ZIP) dosyasına OOXML parçaları elle eklenir:
#
#   xl/pivotCache/pivotCacheDefinition1.xml   alan listesi + benzersiz değerler
#   xl/pivotCache/pivotCacheRecords1.xml      önbellek kayıtları
#   xl/pivotTables/pivotTable1.xml            pivotun kendisi
#   + üç ilişki dosyası, workbook.xml ve [Content_Types].xml güncellemeleri
#
# Önbellek GERÇEK kayıtlarla yazılır; dosya kendi kendine yeterlidir ve
# açılır açılmaz doludur. Ayrıca refreshOnLoad="1" verilir, böylece Excel
# açılışta kaynaktan tazeler.
#
# ----------------------------------------------------------------------------
# EXCEL'İN ÜÇ KATI KURALI  (ihlal edilirse Excel dosyayı AÇARKEN KAPANIR)
# ----------------------------------------------------------------------------
# Pivot, önbellekteki değerlere ADIYLA değil İNDEKSLE işaret eder. Bu yüzden
# üçünün birbirini tutması gerekir:
#
#   1) Eksene (satır/sütun/filtre) konan her alanın önbellekte PAYLAŞILAN
#      ÖĞE LİSTESİ (sharedItems) olmak zorundadır. Listesi olmayan bir alan
#      eksene konursa pivot var olmayan bir indekse bakar.
#   2) Eksene konan her pivotField'ın <items> listesi olmalıdır; uzunluğu
#      paylaşılan öğe sayısı + 1'dir (sondaki "default" öğesi).
#   3) rowItems / colItems, gerçekten kaç satır ve sütun çizileceğini
#      söyler; location/@ref bu sayılarla tutarlı olmalıdır.
#
# Bu kurallar dosya üretildikten sonra _dogrula() içinde tek tek denetlenir.
import datetime
import decimal
import numbers

NS    = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSR   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NSREL = "http://schemas.openxmlformats.org/package/2006/relationships"
NSCT  = "http://schemas.openxmlformats.org/package/2006/content-types"

TIP_CACHE_DEF = ("application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.pivotCacheDefinition+xml")
TIP_CACHE_REC = ("application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.pivotCacheRecords+xml")
TIP_PIVOT     = ("application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.pivotTable+xml")
REL_CACHE_DEF = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                 "relationships/pivotCacheDefinition")
REL_CACHE_REC = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                 "relationships/pivotCacheRecords")
REL_PIVOT     = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                 "relationships/pivotTable")


# XML 1.0'da YASAK olan denetim karakterleri. Eski sistemlerden gelen metin
# alanlarında bulunabilirler; temizlenmezse dosya bozulur.
_GECERSIZ_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]")

# Excel hücresi IEEE double tutar: 2^53'ten büyük tam sayılar TAM olarak
# saklanamaz. Bu sınırın üstündeki bir toplam sayfada yanlış görünür.
EXCEL_TAM_SAYI_SINIRI = 2 ** 53          # 9.007.199.254.740.992
EXCEL_MAKS_SATIR = 1_048_576
EXCEL_MAKS_SUTUN = 16_384


def _oznitelik(v) -> str:
    """XML ÖZNİTELİĞİ için güvenli metin.

    DİKKAT: xml.sax.saxutils.escape yalnız &, < ve > kaçırır — ÇİFT TIRNAĞI
    KAÇIRMAZ. İçinde " geçen bir değer (ör. bir açıklama alanındaki tırnak)
    <s v="C"D"/> gibi BOZUK XML üretir ve dosya hiç açılmaz. Öznitelik
    bağlamında tırnaklar da kaçırılmalıdır.
    """
    return escape(_GECERSIZ_XML.sub("", str(v)), {'"': "&quot;", "'": "&apos;"})


def _sayi_metni(v) -> str:
    """Sayıyı XML'e yazılacak biçime çevirir (tam sayıysa ondalık kısım yok).

    Tam sayılar float'a ÇEVRİLMEDEN yazılır: 8 milyon satırlık bir tablodaki
    SUM sonucu float'ın 15-16 anlamlı basamağını aşabilir ve son basamaklar
    bozulur. numbers.Integral kullanılır çünkü numpy.int64 Python int'in alt
    sınıfı değildir.
    """
    if isinstance(v, numbers.Integral) and not isinstance(v, bool):
        return str(int(v))
    if isinstance(v, decimal.Decimal):
        return format(v.normalize(), "f")   # bilimsel gösterim yok
    f = float(v)
    return str(int(f)) if f.is_integer() else repr(f)


def _sayisal_mi(deger) -> bool:
    # numbers.Number kullanılır: numpy.int64 Python int'in alt sınıfı DEĞİLDİR,
    # isinstance(v, int) ile bakılırsa sayısal kolonlar metin sanılır.
    return isinstance(deger, numbers.Number) and not isinstance(deger, bool)


def _alan_bilgisi(df: pd.DataFrame, boyut_adlari) -> list:
    """Her kolon için önbellek alanı tanımı.

    BOYUT -> paylaşılan öğe listesi yazılır, kayıtlar öğeye İNDEKSLE
             (<x v="i"/>) işaret eder.
    ÖLÇÜ  -> öğe listesi yazılmaz, kayıt sayıyı doğrudan taşır (<n v=".."/>).

    AYRIM ROLE GÖRE YAPILIR, TÜRE GÖRE DEĞİL. Eskiden türe bakılıyordu ve
    COHORT gibi SAYISAL bir boyut "ölçü" muamelesi görüp listesiz kalıyordu;
    pivot onu satır eksenine koyup var olmayan 0. öğeye işaret ediyor, Excel
    de dosyayı açar açmaz kapanıyordu. Rol bilgisi 3. bölümden gelir.
    """
    boyutlar = {str(b) for b in boyut_adlari}
    bilgi = []
    for kolon in df.columns:
        seri = df[kolon]
        bos = bool(seri.isna().any())
        if str(kolon) not in boyutlar:
            bilgi.append({"paylasilan": False, "bos": bos})
            continue

        ogeler, gorulen = [], {}
        for v in seri:
            if pd.isna(v):
                continue
            anahtar = str(v)
            if anahtar in gorulen:
                # ÇAKIŞMA: iki FARKLI değer aynı metne dönüşüyor (ör. 1 ile "1").
                # Sessizce birleştirilirse iki ayrı grup tek öğeye çöker, küpte
                # aynı anahtardan iki satır kalır ve raporda toplamlar tutmaz.
                onceki = gorulen[anahtar]
                if type(onceki) is not type(v):
                    raise ValueError(
                        f"'{kolon}' kolonunda KARIŞIK TÜR var: {onceki!r} "
                        f"({type(onceki).__name__}) ile {v!r} ({type(v).__name__}) "
                        f"aynı metne ('{anahtar}') dönüşüyor.\n"
                        "        Boyut kolonu tek tipte olmalı; kaynakta bu kolonu "
                        "TO_CHAR/CAST ile tek tipe çevirin\n"
                        "        ya da ALAN_HARIC listesine ekleyin.")
                continue
            gorulen[anahtar] = v
            ogeler.append(v)

        if not ogeler:
            tip = "metin"
        elif all(_sayisal_mi(o) for o in ogeler):
            tip = "sayi"
        elif all(isinstance(o, (datetime.datetime, datetime.date)) for o in ogeler):
            # Tarihler METİN yazılırsa hem etiketler "2026-04-01 00:00:00" gibi
            # görünür hem de refreshOnLoad ile sayfadan tazelendiğinde sayfadaki
            # gerçek tarihlerle çelişir. <d> öğesi kullanılır.
            tip = "tarih"
        else:
            tip = "metin"
        bilgi.append({"paylasilan": True, "ogeler": ogeler, "bos": bos,
                      "tip": tip, "sayisal": tip == "sayi"})
    return bilgi


def _oge_sayisi(alan) -> int:
    """Alanın paylaşılan öğe sayısı (boş değer de bir öğedir)."""
    if not alan.get("paylasilan"):
        return 0
    return len(alan["ogeler"]) + (1 if alan["bos"] else 0)


def _paylasilan_xml(alan) -> str:
    if not alan["paylasilan"]:
        ozn = ['containsSemiMixedTypes="0"', 'containsString="0"', 'containsNumber="1"']
        if alan["bos"]:
            ozn.append('containsBlank="1"')
        return f'<sharedItems {" ".join(ozn)}/>'

    ogeler, bos = alan["ogeler"], alan["bos"]
    if not ogeler and not bos:
        return "<sharedItems/>"

    ozn = []
    if bos:
        ozn.append('containsBlank="1"')
    if alan.get("tip") == "tarih":
        tarihler = [pd.Timestamp(o).to_pydatetime() for o in ogeler]
        ozn += ['containsSemiMixedTypes="0"', 'containsString="0"',
                'containsNonDate="0"', 'containsDate="1"',
                f'minDate="{min(tarihler).isoformat()}"',
                f'maxDate="{max(tarihler).isoformat()}"']
        icerik = "".join(f'<d v="{pd.Timestamp(o).to_pydatetime().isoformat()}"/>'
                         for o in ogeler)
    elif alan["sayisal"]:
        sayilar = [float(o) for o in ogeler]
        ozn += ['containsSemiMixedTypes="0"', 'containsString="0"', 'containsNumber="1"']
        if all(s.is_integer() for s in sayilar):
            ozn.append('containsInteger="1"')
        ozn += [f'minValue="{_sayi_metni(min(sayilar))}"',
                f'maxValue="{_sayi_metni(max(sayilar))}"']
        icerik = "".join(f'<n v="{_sayi_metni(o)}"/>' for o in ogeler)
    else:
        icerik = "".join(f'<s v="{_oznitelik(o)}"/>' for o in ogeler)
    # Boş değer listenin SONUNA eklenir; kayıtlardaki indeks de buna göre verilir.
    if bos:
        icerik += "<m/>"
    ozn.append(f'count="{_oge_sayisi(alan)}"')
    return f'<sharedItems {" ".join(ozn)}>{icerik}</sharedItems>'


def _cache_def_xml(ref, veri_sayfasi, kolonlar, bilgi, kayit_sayisi):
    alanlar = [f'<cacheField name="{_oznitelik(kolon)}" numFmtId="0">'
               f'{_paylasilan_xml(alan)}</cacheField>'
               for kolon, alan in zip(kolonlar, bilgi)]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<pivotCacheDefinition xmlns="{NS}" xmlns:r="{NSR}" r:id="rId1" '
            f'refreshOnLoad="1" refreshedVersion="8" minRefreshableVersion="3" '
            f'createdVersion="8" recordCount="{kayit_sayisi}">'
            f'<cacheSource type="worksheet">'
            f'<worksheetSource ref="{ref}" sheet="{_oznitelik(veri_sayfasi)}"/></cacheSource>'
            f'<cacheFields count="{len(kolonlar)}">{"".join(alanlar)}</cacheFields>'
            '</pivotCacheDefinition>')


def _cache_rec_xml(df: pd.DataFrame, bilgi):
    haritalar = []
    for alan in bilgi:
        if alan["paylasilan"]:
            haritalar.append(({str(o): i for i, o in enumerate(alan["ogeler"])},
                              len(alan["ogeler"]) if alan["bos"] else None))
        else:
            haritalar.append(None)

    satirlar = []
    for satir in df.itertuples(index=False, name=None):
        parca = []
        for deger, alan, harita in zip(satir, bilgi, haritalar):
            if alan["paylasilan"]:
                parca.append(f'<x v="{harita[1] if pd.isna(deger) else harita[0][str(deger)]}"/>')
            elif pd.isna(deger):
                parca.append("<m/>")
            elif _sayisal_mi(deger):
                parca.append(f'<n v="{_sayi_metni(deger)}"/>')
            else:
                # Ölçü kolonunda sayı olmayan bir değer var: sessizce yanlış
                # bir dosya üretmek yerine sebebi söylenir.
                raise ValueError(
                    f"Ölçü kolonunda sayı olmayan değer: {deger!r}\n"
                    "        OLCULER sadece count/sum üretir; bu kolon boyut olmalı.")
        satirlar.append(f'<r>{"".join(parca)}</r>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<pivotCacheRecords xmlns="{NS}" xmlns:r="{NSR}" count="{len(satirlar)}">'
            f'{"".join(satirlar)}</pivotCacheRecords>')


def _eksen_ogeleri(n: int, genel_toplam: bool) -> list:
    """rowItems / colItems girdileri: her öğe için bir <i>, sonda Genel Toplam."""
    parca = ["<i><x/></i>" if k == 0 else f'<i><x v="{k}"/></i>' for k in range(n)]
    if genel_toplam:
        parca.append('<i t="grand"><x/></i>')
    return parca


def _pivot_xml(kolonlar, bilgi, satir_idx, sutun_idx, deger_idx):
    alanlar = []
    for i in range(len(kolonlar)):
        if i in (satir_idx, sutun_idx) and i is not None:
            eksen = "axisRow" if i == satir_idx else "axisCol"
            n = _oge_sayisi(bilgi[i])
            ogeler = "".join(f'<item x="{k}"/>' for k in range(n)) + '<item t="default"/>'
            alanlar.append(f'<pivotField axis="{eksen}" showAll="0">'
                           f'<items count="{n + 1}">{ogeler}</items></pivotField>')
        elif i == deger_idx:
            alanlar.append('<pivotField dataField="1" showAll="0"/>')
        else:
            alanlar.append('<pivotField showAll="0"/>')

    deger_var = deger_idx is not None

    if satir_idx is not None:
        satir_ogeleri = _eksen_ogeleri(_oge_sayisi(bilgi[satir_idx]), deger_var)
        satir_alani = (f'<rowFields count="1"><field x="{satir_idx}"/></rowFields>'
                       f'<rowItems count="{len(satir_ogeleri)}">'
                       f'{"".join(satir_ogeleri)}</rowItems>')
        satir_yuksekligi = len(satir_ogeleri)
    else:
        satir_alani = '<rowItems count="1"><i/></rowItems>'
        satir_yuksekligi = 1

    if sutun_idx is not None:
        sutun_ogeleri = _eksen_ogeleri(_oge_sayisi(bilgi[sutun_idx]), deger_var)
        sutun_alani = (f'<colFields count="1"><field x="{sutun_idx}"/></colFields>'
                       f'<colItems count="{len(sutun_ogeleri)}">'
                       f'{"".join(sutun_ogeleri)}</colItems>')
        sutun_genisligi = len(sutun_ogeleri)
    else:
        sutun_alani = '<colItems count="1"><i/></colItems>'
        sutun_genisligi = 1

    deger_alani = ""
    if deger_var:
        deger_alani = (f'<dataFields count="1"><dataField name="Toplam '
                       f'{_oznitelik(kolonlar[deger_idx])}" fld="{deger_idx}" '
                       f'baseField="0" baseItem="0"/></dataFields>')

    # Yerleşim A3'ten başlar. Sütun alanı varsa iki başlık satırı olur
    # (üstte ölçü/sütun alanı adı, altta sütun etiketleri); yoksa bir.
    baslik_satiri = 2 if sutun_idx is not None else 1
    toplam_satir = baslik_satiri + satir_yuksekligi
    toplam_sutun = 1 + sutun_genisligi
    ref = f"A3:{get_column_letter(max(2, toplam_sutun))}{2 + toplam_satir}"

    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<pivotTableDefinition xmlns="{NS}" name="PivotTable1" cacheId="1" '
            'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" '
            'applyPatternFormats="0" applyAlignmentFormats="0" '
            'applyWidthHeightFormats="1" dataCaption="Değerler" updatedVersion="8" '
            'minRefreshableVersion="3" useAutoFormatting="1" itemPrintTitles="1" '
            'createdVersion="8" indent="0" outline="1" outlineData="1" '
            'multipleFieldFilters="0">'
            f'<location ref="{ref}" firstHeaderRow="1" '
            f'firstDataRow="{2 if sutun_idx is not None else 1}" firstDataCol="1"/>'
            f'<pivotFields count="{len(kolonlar)}">{"".join(alanlar)}</pivotFields>'
            f'{satir_alani}{sutun_alani}{deger_alani}'
            '<pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" '
            'showColHeaders="1" showRowStripes="0" showColStripes="0" '
            'showLastColumn="1"/>'
            '</pivotTableDefinition>')


# ============================================================================
# 6. PARÇALARI .xlsx (ZIP) İÇİNE YERLEŞTİRME
# ============================================================================
def _enjekte(yol: Path, df: pd.DataFrame, ref: str, boyut_adlari,
             satir_idx, sutun_idx, deger_idx) -> None:
    """Pivot parçalarını hazır bir .xlsx dosyasına ekler."""
    kolonlar = [str(c) for c in df.columns]
    # Boyut listesi ROL bilgisidir; hangi alanın paylaşılan öğe listesine
    # ihtiyacı olduğunu bu belirler (bkz. 5. bölümdeki üç kat kuralı).
    bilgi = _alan_bilgisi(df, boyut_adlari)

    with zipfile.ZipFile(yol) as zin:
        icerik = {a: zin.read(a) for a in zin.namelist()}

    # --- pivot sayfasının dosya yolunu bul --------------------------------
    # Öznitelik SIRASINDAN bağımsız olsun diye XML ayrıştırılır; regex ile
    # aranırsa openpyxl'in yazdığı sırada (Target önce, Id sonra) bulunamaz.
    wb_xml = icerik["xl/workbook.xml"].decode("utf-8")
    rels_xml = icerik["xl/_rels/workbook.xml.rels"].decode("utf-8")

    rid = next((sh.get(f"{{{NSR}}}id") for sh in ET.fromstring(wb_xml).iter(f"{{{NS}}}sheet")
                if sh.get("name") == PIVOT_SAYFASI), None)
    if rid is None:
        raise RuntimeError(f"'{PIVOT_SAYFASI}' sayfası workbook.xml içinde bulunamadı.")
    hedef = next((r.get("Target") for r in ET.fromstring(rels_xml).iter(f"{{{NSREL}}}Relationship")
                  if r.get("Id") == rid), None)
    if hedef is None:
        raise RuntimeError(f"{rid} ilişkisi workbook.xml.rels içinde bulunamadı.")
    hedef = hedef.lstrip("/")
    sheet_yolu = hedef if hedef.startswith("xl/") else "xl/" + hedef

    # --- yeni parçalar ----------------------------------------------------
    icerik["xl/pivotCache/pivotCacheDefinition1.xml"] = _cache_def_xml(
        ref, VERI_SAYFASI, kolonlar, bilgi, len(df)).encode("utf-8")
    icerik["xl/pivotCache/pivotCacheRecords1.xml"] = _cache_rec_xml(df, bilgi).encode("utf-8")

    # Satır/sütun sayısı ARTIK TAHMİN EDİLMİYOR: yerleşim, önbelleğe gerçekten
    # yazılan öğe sayısından üretilir. Eskiden nunique() ile tahmin ediliyor ve
    # rowItems tek satır diyorken location altı satır iddia ediyordu.
    icerik["xl/pivotTables/pivotTable1.xml"] = _pivot_xml(
        kolonlar, bilgi, satir_idx, sutun_idx, deger_idx).encode("utf-8")

    icerik["xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{NSREL}"><Relationship Id="rId1" Type="{REL_CACHE_REC}" '
        'Target="pivotCacheRecords1.xml"/></Relationships>').encode("utf-8")
    icerik["xl/pivotTables/_rels/pivotTable1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<Relationships xmlns="{NSREL}"><Relationship Id="rId1" Type="{REL_CACHE_DEF}" '
        'Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>').encode("utf-8")

    # --- workbook.xml: <pivotCaches> --------------------------------------
    # xmlns:r BURADA tanımlanır: openpyxl bu öneki <workbook> kökünde değil,
    # her <sheet> öğesinde ayrı ayrı tanımlıyor. Kökte var sanıp r:id yazarsak
    # XML "unbound prefix" ile bozulur ve Excel dosyayı "onarılması gerekiyor"
    # diye açar.
    if "<pivotCaches>" not in wb_xml:
        wb_xml = wb_xml.replace(
            "</workbook>",
            f'<pivotCaches xmlns:r="{NSR}">'
            '<pivotCache cacheId="1" r:id="rId900"/></pivotCaches></workbook>')
    icerik["xl/workbook.xml"] = wb_xml.encode("utf-8")
    icerik["xl/_rels/workbook.xml.rels"] = rels_xml.replace(
        "</Relationships>",
        f'<Relationship Id="rId900" Type="{REL_CACHE_DEF}" '
        'Target="pivotCache/pivotCacheDefinition1.xml"/></Relationships>').encode("utf-8")

    # --- sayfa ilişkisi ---------------------------------------------------
    sheet_rels = sheet_yolu.replace("worksheets/", "worksheets/_rels/") + ".rels"
    if sheet_rels in icerik:
        s = icerik[sheet_rels].decode("utf-8").replace(
            "</Relationships>",
            f'<Relationship Id="rId1000" Type="{REL_PIVOT}" '
            'Target="../pivotTables/pivotTable1.xml"/></Relationships>')
    else:
        s = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             f'<Relationships xmlns="{NSREL}"><Relationship Id="rId1000" Type="{REL_PIVOT}" '
             'Target="../pivotTables/pivotTable1.xml"/></Relationships>')
    icerik[sheet_rels] = s.encode("utf-8")

    # --- content types ----------------------------------------------------
    ctx = icerik["[Content_Types].xml"].decode("utf-8")
    ctx = ctx.replace("</Types>",
        f'<Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="{TIP_CACHE_DEF}"/>'
        f'<Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="{TIP_CACHE_REC}"/>'
        f'<Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="{TIP_PIVOT}"/>'
        "</Types>")
    icerik["[Content_Types].xml"] = ctx.encode("utf-8")

    gecici = yol.with_suffix(".tmp.xlsx")
    with zipfile.ZipFile(gecici, "w", zipfile.ZIP_DEFLATED) as zout:
        for ad, veri in icerik.items():
            zout.writestr(ad, veri)
    shutil.move(str(gecici), str(yol))


def _dogrula(yol: Path) -> list:
    """Üretilen dosyanın OOXML bütünlüğünü kontrol eder (Excel açmadan).

    Kırık bir ilişki ya da eksik content-type, Excel'de "dosya onarılmalı"
    uyarısına yol açar. Bu kontrol onu üretim anında yakalar.
    """
    sorunlar = []
    z = zipfile.ZipFile(yol)
    adlar = set(z.namelist())

    for a in sorted(adlar):
        if a.endswith((".xml", ".rels")):
            try:
                ET.fromstring(z.read(a))
            except ET.ParseError as e:
                sorunlar.append(f"XML bozuk: {a} -> {e}")

    for a in sorted(adlar):
        if not a.endswith(".rels"):
            continue
        taban = dirname(dirname(a)) or ""
        for rel in ET.fromstring(z.read(a)).iter(f"{{{NSREL}}}Relationship"):
            if rel.get("TargetMode") == "External":
                continue
            t = rel.get("Target", "")
            coz = t.lstrip("/") if t.startswith("/") else normpath(join(taban, t))
            if coz not in adlar:
                sorunlar.append(f"kırık ilişki: {a} [{rel.get('Id')}] -> {t}")

    _pivot_tutarlilik(z, sorunlar)

    ctx = ET.fromstring(z.read("[Content_Types].xml"))
    varsayilan = {d.get("Extension", "").lower() for d in ctx.iter(f"{{{NSCT}}}Default")}
    ozel = {o.get("PartName", "").lstrip("/"): (o.get("ContentType") or "")
            for o in ctx.iter(f"{{{NSCT}}}Override")}

    # Kayıtlı olması yetmez, DOĞRU tip olmalı: yanlış ya da boş bir
    # content-type Excel'de "dosya onarılmalı" uyarısına yol açar.
    for parca, beklenen in (("xl/pivotCache/pivotCacheDefinition1.xml", TIP_CACHE_DEF),
                            ("xl/pivotCache/pivotCacheRecords1.xml", TIP_CACHE_REC),
                            ("xl/pivotTables/pivotTable1.xml", TIP_PIVOT)):
        if ozel.get(parca) != beklenen:
            sorunlar.append(f"content-type yanlış: {parca} -> {ozel.get(parca)!r}, "
                            f"beklenen {beklenen!r}")
    for a in sorted(adlar):
        if a == "[Content_Types].xml" or a in ozel:
            continue
        if (a.rsplit(".", 1)[-1].lower() if "." in a else "") not in varsayilan:
            sorunlar.append(f"content-type YOK: {a}")

    return sorunlar


def _pivot_tutarlilik(z, sorunlar: list) -> None:
    """Excel'in üç katı kuralını denetler (bkz. 5. bölüm).

    Bu kontrol, Excel'i AÇMADAN "dosya açılır açılmaz kapanır mı" sorusunu
    cevaplar. Bir öğe indeksi listenin dışına taşarsa Excel onarım uyarısı
    bile vermeden kapanır; burada yakalanır.
    """
    try:
        cd = ET.fromstring(z.read("xl/pivotCache/pivotCacheDefinition1.xml"))
        rec = ET.fromstring(z.read("xl/pivotCache/pivotCacheRecords1.xml"))
        pt = ET.fromstring(z.read("xl/pivotTables/pivotTable1.xml"))
    except KeyError as e:
        sorunlar.append(f"pivot parçası eksik: {e}")
        return

    # --- her alanın paylaşılan öğe sayısı --------------------------------
    adlar, sayilar = [], []
    for cf in cd.iter(f"{{{NS}}}cacheField"):
        adlar.append(cf.get("name"))
        si = cf.find(f"{{{NS}}}sharedItems")
        cocuk = list(si) if si is not None else []
        sayilar.append(len(cocuk))
        if si is not None and si.get("count") and int(si.get("count")) != len(cocuk):
            sorunlar.append(f"'{cf.get('name')}' sharedItems count={si.get('count')} "
                            f"ama {len(cocuk)} öğe var")

    # Aynı adlı iki alan: Excel önbelleği kabul etmez.
    tekrar = sorted({a for a in adlar if adlar.count(a) > 1})
    if tekrar:
        sorunlar.append(f"AYNI ADLI alan(lar): {tekrar} — sorguda kolon adlarını "
                        "AS ile ayırın")
    if any(not a for a in adlar):
        sorunlar.append("adı boş alan var")

    cfs = cd.find(f"{{{NS}}}cacheFields")
    if cfs.get("count") and int(cfs.get("count")) != len(adlar):
        sorunlar.append(f"cacheFields count={cfs.get('count')} ama {len(adlar)} alan var")
    kayit_sayisi = len(list(rec.iter(f"{{{NS}}}r")))
    if cd.get("recordCount") and int(cd.get("recordCount")) != kayit_sayisi:
        sorunlar.append(f"recordCount={cd.get('recordCount')} ama {kayit_sayisi} kayıt var")
    if rec.get("count") and int(rec.get("count")) != kayit_sayisi:
        sorunlar.append(f"pivotCacheRecords count={rec.get('count')} ama "
                        f"{kayit_sayisi} kayıt var")

    # Kaynak aralık, önbellekte duran kayıt sayısı ve alan sayısıyla tutmalı:
    # refreshOnLoad="1" olduğu için Excel açılışta bu aralıktan tazeler.
    ws = cd.find(f"{{{NS}}}cacheSource/{{{NS}}}worksheetSource")
    if ws is not None:
        try:
            bas, _, son = ws.get("ref", "").partition(":")
            ilk_s = int(re.sub(r"[A-Z]", "", bas)); son_s = int(re.sub(r"[A-Z]", "", son))
            if son_s - ilk_s != kayit_sayisi:
                sorunlar.append(f"kaynak aralık {son_s - ilk_s} veri satırı gösteriyor, "
                                f"önbellekte {kayit_sayisi} kayıt var")
            if son_s > EXCEL_MAKS_SATIR:
                sorunlar.append(f"kaynak aralık {son_s}. satıra taşıyor "
                                f"(Excel sınırı {EXCEL_MAKS_SATIR})")
        except (ValueError, TypeError):
            sorunlar.append(f"kaynak aralık okunamadı: {ws.get('ref')!r}")

    # --- kayıtlar: her <x v="i"/> listenin içinde mi? --------------------
    for no, r in enumerate(rec.iter(f"{{{NS}}}r"), start=1):
        if len(list(r)) != len(adlar):
            sorunlar.append(f"kayıt {no}: {len(list(r))} değer var, {len(adlar)} alan bekleniyor")
            return
        for i, oge in enumerate(list(r)):
            if not oge.tag.endswith("}x"):
                continue
            idx = int(oge.get("v", "0"))
            if i >= len(sayilar) or idx >= sayilar[i]:
                sorunlar.append(
                    f"kayıt {no}: '{adlar[i] if i < len(adlar) else i}' alanında "
                    f"indeks {idx}, oysa listede {sayilar[i] if i < len(sayilar) else 0} öğe var")
                return  # aynı hata her satırda tekrarlar, bir kez söylemek yeter

    # --- eksene konan alanlar --------------------------------------------
    pf = list(pt.find(f"{{{NS}}}pivotFields"))
    if len(pf) != len(sayilar):
        # Devam edilirse indeksler kayar ve denetim IndexError ile patlar;
        # sebebi burada net söylenip durulur.
        sorunlar.append(f"pivotField sayısı {len(pf)}, cacheField sayısı {len(sayilar)}")
        return

    for etiket, alan_etiketi, oge_etiketi in (("satır", "rowFields", "rowItems"),
                                              ("sütun", "colFields", "colItems")):
        alanlar = pt.find(f"{{{NS}}}{alan_etiketi}")
        if alanlar is None:
            continue
        idx = int(list(alanlar)[0].get("x"))
        if idx < 0 or idx >= len(sayilar):
            sorunlar.append(f"{etiket} alanı indeksi {idx} alan listesinin dışında")
            continue
        if sayilar[idx] == 0:
            sorunlar.append(
                f"{etiket} ekseninde '{adlar[idx]}' var ama önbellekte paylaşılan "
                "öğe listesi YOK — Excel açılırken kapanır")
        ogeler = pf[idx].find(f"{{{NS}}}items")
        if ogeler is None:
            sorunlar.append(f"'{adlar[idx]}' eksene konmuş ama <items> listesi yok")
        elif len(list(ogeler)) != sayilar[idx] + 1:
            sorunlar.append(f"'{adlar[idx]}' <items> uzunluğu {len(list(ogeler))}, "
                            f"beklenen {sayilar[idx] + 1}")

        cizilen = pt.find(f"{{{NS}}}{oge_etiketi}")
        beklenen = sayilar[idx] + (1 if pt.find(f"{{{NS}}}dataFields") is not None else 0)
        if cizilen is not None and len(list(cizilen)) != beklenen:
            sorunlar.append(f"{oge_etiketi}: {len(list(cizilen))} girdi, beklenen {beklenen}")

    # --- değer alanı gerçekten ölçüye mi bağlı? --------------------------
    for dfl in pt.iter(f"{{{NS}}}dataField"):
        fld = int(dfl.get("fld", "-1"))
        if not (0 <= fld < len(adlar)):
            sorunlar.append(f"dataField fld={fld} alan listesinin dışında")
        elif pf[fld].get("dataField") != "1":
            sorunlar.append(f"dataField '{adlar[fld]}' gösteriyor ama o pivotField "
                            'üzerinde dataField="1" yok')

    # --- location/@ref gerçekten çizilecek alanı kapsıyor mu? ------------
    loc = pt.find(f"{{{NS}}}location")
    satirlar = pt.find(f"{{{NS}}}rowItems")
    sutunlar = pt.find(f"{{{NS}}}colItems")
    if loc is not None and satirlar is not None and sutunlar is not None:
        bas, son = loc.get("ref", "A1:A1").split(":")
        ilk_satir = int(re.sub(r"[A-Z]", "", bas))
        son_satir = int(re.sub(r"[A-Z]", "", son))
        baslik = 2 if pt.find(f"{{{NS}}}colFields") is not None else 1
        beklenen = baslik + len(list(satirlar))
        if son_satir - ilk_satir + 1 != beklenen:
            sorunlar.append(f"location satır sayısı {son_satir - ilk_satir + 1}, "
                            f"rowItems'a göre beklenen {beklenen}")
        if son_satir > EXCEL_MAKS_SATIR:
            sorunlar.append(f"pivot {son_satir}. satıra taşıyor "
                            f"(Excel sınırı {EXCEL_MAKS_SATIR})")


# ============================================================================
# 7. RAPOR ÜRETİMİ
# ============================================================================
def pivot_excel_uret(cikti: Path = CIKTI_DOSYASI) -> Path:
    """Özet küpü çeker, Excel'i yazar ve içine gerçek PivotTable yerleştirir."""
    print("=" * 78)
    print(f"PİVOTTABLE RAPORU  (veri kaynağı: {VERI_KAYNAGI})")
    print("=" * 78)

    # Hangi kolon boyut, hangisi ölçü: elle yazılmaz, veriden okunur.
    boyutlar, olculer = alanlari_belirle()
    satir_ad, sutun_ad, deger_ad = _acilis_yerlesimi(boyutlar, olculer)

    df = kup_verisi()
    if df.empty:
        raise RuntimeError("Özet küp BOŞ — sorgu hiç satır döndürmedi.")

    kolonlar = [str(c) for c in df.columns]
    olcu_adlari = [o["ad"] for o in olculer]

    def _indeks(ad, nerede):
        if ad is None:
            return None
        if ad not in kolonlar:
            raise ValueError(f"{nerede} = {ad!r} küpte yok. Kullanılabilir alanlar: {kolonlar}")
        return kolonlar.index(ad)

    satir_idx = _indeks(satir_ad, "PIVOT_SATIR")
    sutun_idx = _indeks(sutun_ad, "PIVOT_SUTUN")
    deger_idx = _indeks(deger_ad, "PIVOT_DEGER")

    # Eksene SADECE boyut konabilir. Ölçünün paylaşılan öğe listesi yoktur;
    # eksene konursa Excel var olmayan bir öğeye bakar ve dosya açılır
    # açılmaz kapanır. Sessizce bozuk dosya üretmek yerine burada durulur.
    for ad, nerede in ((satir_ad, "PIVOT_SATIR"), (sutun_ad, "PIVOT_SUTUN")):
        if ad is not None and ad not in boyutlar:
            raise ValueError(
                f"{nerede} = {ad!r} bir BOYUT değil (boyutlar: {boyutlar}).\n"
                "        Satır ve sütun eksenine sadece boyut konabilir; ölçüler "
                "'Değerler' alanına gider.")
    if deger_idx is not None and deger_ad not in olcu_adlari:
        print(f"[UYARI] PIVOT_DEGER = {deger_ad!r} bir ÖLÇÜ değil ({olcu_adlari}). "
              "Excel bu alanı yine de toplayacaktır, ama anlamlı olmayabilir.")

    df = _metni_temizle(df)
    _olculeri_dogrula()
    _excel_sinir_kontrolu(df)
    _hassasiyet_uyarisi(df, boyutlar)

    # --- 1) GEÇİCİ dosyaya üret ------------------------------------------
    # Hedef dosya, HER ŞEY DOĞRULANANA KADAR hiç yazılmaz. Sebebi: Excel'i
    # çökerten bir dosya, Excel'in "devre dışı bırakılan öğeler" listesine
    # düşüp sonraki açılışlarda Güvenli Mod'a sebep oluyor. Ortada yarım
    # kalmış bir dosya bırakmamak, bunun tek kesin çaresi.
    #
    # Yan fayda: üretim yarıda kalırsa önceki SAĞLAM çıktı yerinde kalır.
    cikti = Path(cikti)
    gecici = cikti.with_name(cikti.stem + ".uretiliyor.xlsx")
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = VERI_SAYFASI
        ws.append(kolonlar)
        for hucre in ws[1]:
            hucre.font = Font(bold=True, color="FFFFFF")
            hucre.fill = PatternFill("solid", fgColor="14365C")
        for satir in df.itertuples(index=False, name=None):
            ws.append([None if pd.isna(v) else v for v in satir])
        ws.freeze_panes = "A2"
        for i, kolon in enumerate(kolonlar, start=1):
            ws.column_dimensions[get_column_letter(i)].width = max(12, min(30, len(kolon) + 4))
        if VERI_SAYFASI_GIZLE:
            ws.sheet_state = "hidden"

        wp = wb.create_sheet(PIVOT_SAYFASI)
        wp["A1"] = "PivotTable"
        wp["A1"].font = Font(bold=True, size=13, color="14365C")
        wp["A2"] = ("Alanları sürükleyerek satır / sütun / filtre / değer yerleşimini "
                    "değiştirebilirsiniz.")
        wp["A2"].font = Font(size=9, color="7C8B9C")
        wb.move_sheet(PIVOT_SAYFASI, offset=-1)   # pivot sayfası önde açılsın
        wb.active = 0
        wb.save(gecici)

        # --- 2) pivot parçalarını yerleştir -------------------------------
        ref = f"A1:{get_column_letter(len(kolonlar))}{len(df) + 1}"
        _enjekte(gecici, df, ref, boyutlar, satir_idx, sutun_idx, deger_idx)

        # --- 3) doğrula ----------------------------------------------------
        try:
            sorunlar = _dogrula(gecici)
        except Exception as e:
            raise RuntimeError(f"Dosya denetlenemedi. Sebep: {type(e).__name__}: {e}") from e
        if sorunlar:
            raise RuntimeError(
                f"Dosya bütünlük kontrolünde {len(sorunlar)} sorun bulundu:\n        - "
                + "\n        - ".join(sorunlar))
        print("[BİLGİ] Dosya bütünlük kontrolü GEÇTİ "
              "(XML, ilişkiler, content-type'lar ve pivot indeksleri tutarlı).")
    except BaseException:
        # Hangi sebeple olursa olsun: yarım kalmış dosya diskte BIRAKILMAZ.
        gecici.unlink(missing_ok=True)
        raise

    # --- 4) ancak şimdi yerine koy ---------------------------------------
    shutil.move(str(gecici), str(cikti))

    print("\n" + "=" * 78)
    print(f"RAPOR OLUŞTURULDU: {cikti.resolve()}")
    print(f"  Özet küp        : {len(df):,} satır x {len(kolonlar)} kolon"
          .replace(",", "."))
    print(f"  Pivot alanları  : {kolonlar}")
    print(f"  Açılış yerleşimi: satır={satir_ad} · sütun={sutun_ad} · değer={deger_ad}")
    print(f"  Sayfalar        : {[s.title for s in Workbook().worksheets] if False else 'Pivot, ' + VERI_SAYFASI}")
    print("  Excel'de PivotTable'ı tıklayıp alanları sürükleyebilirsiniz.")
    print("=" * 78)
    return cikti


def _excel_sinir_kontrolu(df: pd.DataFrame) -> None:
    """Küp Excel'in sayfa sınırlarına sığıyor mu?

    Sığmazsa dosya yazılabilir ama Excel açamaz. Sınırı burada, dosya
    üretilmeden önce ve sebebini söyleyerek kesmek gerekir.
    """
    gerekli = len(df) + 1                      # +1: başlık satırı
    if gerekli > EXCEL_MAKS_SATIR:
        raise RuntimeError(
            f"Özet küp {len(df):,} satır — Excel sayfasına SIĞMAZ "
            f"(sınır {EXCEL_MAKS_SATIR:,} satır).".replace(",", ".") + "\n"
            "        Boyut sayısını azaltın (MAKS_BOYUT_SAYISI) ya da yüksek "
            "kardinaliteli bir boyutu ALAN_HARIC'e ekleyin.")
    if len(df.columns) > EXCEL_MAKS_SUTUN:
        raise RuntimeError(f"Küpte {len(df.columns)} kolon var, Excel sınırı "
                           f"{EXCEL_MAKS_SUTUN}.")


def _hassasiyet_uyarisi(df: pd.DataFrame, boyut_adlari) -> None:
    """Excel hücresinde TAM saklanamayacak kadar büyük sayı var mı?

    Excel hücresi IEEE double tutar: 2^53'ten (9.007.199.254.740.992) büyük
    tam sayılar son basamaklarını kaybeder. Önbelleğe doğru değeri yazsak
    bile refreshOnLoad ile sayfadan tazelendiğinde Excel'in gördüğü sayı
    farklı olur. Bu bir kod hatası değil, Excel'in sınırıdır — ama bir banka
    raporunda sessiz kalmamalı.
    """
    boyutlar = {str(b) for b in boyut_adlari}
    tasan = []
    for kolon in df.columns:
        if str(kolon) in boyutlar:
            continue
        seri = pd.to_numeric(df[kolon], errors="coerce").abs()
        en_buyuk = _tek_deger(seri.max(), f"'{kolon}' en büyük değeri")
        if seri.notna().any() and float(en_buyuk) > EXCEL_TAM_SAYI_SINIRI:
            tasan.append((str(kolon), float(en_buyuk)))
    if not tasan:
        return
    print(f"[UYARI] Excel hücresi en çok {EXCEL_TAM_SAYI_SINIRI:,} büyüklüğünde bir "
          "tam sayıyı TAM saklayabilir.".replace(",", "."))
    for ad, en_buyuk in tasan:
        print(f"        '{ad}' bu sınırı aşıyor (en büyük: {en_buyuk:,.0f}) — "
              "sayfada son basamaklar yanlış görünür.".replace(",", "."))
    print("        Çözüm: ölçüyü kaynakta küçültün (ör. SUM(TUTAR)/1000) ya da "
          "kırılımı daraltın.")

def _metni_temizle(df: pd.DataFrame) -> pd.DataFrame:
    """Excel'in KABUL ETMEDİĞİ denetim karakterlerini metin değerlerden atar.

    XML 1.0 bu karakterleri yasaklar; openpyxl da hücreye yazmayı reddeder
    ("IllegalCharacterError"). Eski sistemlerden gelen açıklama alanlarında
    bulunabilirler.

    Temizlik MERKEZÎ yapılır, çünkü önbellek (_oznitelik) zaten bu
    karakterleri atıyor: sayfa temizlenmezse önbellek ile sayfa BİRBİRİNDEN
    FARKLI olur ve refreshOnLoad ile tazelendiğinde değerler değişir.

    Silinen karakter Excel'de zaten gösterilemez, ama kaç hücrenin
    değiştiği sessiz kalmasın diye yazılır.
    """
    temiz = df.copy()
    degisen = 0
    for kolon in temiz.columns:
        # dtype'a BAKILMAZ: pandas sürümüne göre metin kolonunun dtype'ı
        # 'object' de olabilir 'str' de. 'dtype != object' yazılırsa yeni
        # pandas'ta metin kolonları atlanır ve temizlik hiç çalışmaz.
        if (pd.api.types.is_numeric_dtype(temiz[kolon])
                or pd.api.types.is_datetime64_any_dtype(temiz[kolon])
                or pd.api.types.is_bool_dtype(temiz[kolon])):
            continue
        eski = temiz[kolon]
        yeni = eski.map(lambda v: _GECERSIZ_XML.sub("", v) if isinstance(v, str) else v)
        degisen += int(_tek_deger((eski.astype(str) != yeni.astype(str)).sum(),
                                  f"'{kolon}' değişen hücre sayısı"))
        temiz[kolon] = yeni
    if degisen:
        print(f"[UYARI] {degisen} hücrede Excel'in kabul etmediği denetim karakteri "
              "vardı, temizlendi.")
        print("        Bu karakterler Excel'de zaten gösterilemez; kaynak alanda "
              "eski sistemden kalmış olabilir.")
    return temiz

def _olculeri_dogrula() -> None:
    """OLCULER içindeki işlemler desteklenen işlemler mi?

    Bu kontrol eskiden yalnız Oracle yolundaki sorgu üretiminde vardı; Excel
    yolunda pandas 'mean' gibi bir işlemi sessizce hesaplıyordu. PivotTable
    değerleri YENİDEN TOPLAYARAK çalıştığı için ortalama yeniden toplanamaz:
    dosya açılır ama gösterdiği sayılar yanlış olur. İki yolda da durulur.
    """
    for o in OLCULER:
        islem = o.get("islem", "count")
        if islem not in ISLEM_SQL:
            raise ValueError(
                f"{o.get('ad')!r}: '{islem}' desteklenmiyor. Geçerli: {sorted(ISLEM_SQL)}.\n"
                "        PivotTable değerleri yeniden TOPLAYARAK özetler; ortalama "
                "yeniden toplanamaz.\n"
                "        Ortalama için SUM ve COUNT'u ayrı ölçü olarak ekleyip "
                "Excel'de Hesaplanmış Alan kullanın.")


if __name__ == "__main__":
    pivot_excel_uret()