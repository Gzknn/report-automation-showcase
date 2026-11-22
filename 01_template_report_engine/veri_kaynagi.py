"""Veri kaynağı sınırı: Oracle ya da Excel -> DataFrame.

Motor verinin nereden geldiğini bilmez, yalnızca DataFrame görür. **Otomatik
geçiş yoktur**: Oracle'da sorun çıkarsa hata açıkça verilir, sessizce
Excel'e düşülmez; hangi veriyle çalıştığınız her zaman nettir.

Bağlantı bilgileri KODA YAZILMAZ, ortam değişkenlerinden okunur:
    ORACLE_USER, ORACLE_PASSWORD, ORACLE_DSN   (ör. "sunucu:port/servis")
Kurulum (yalnız Oracle için):  pip install oracledb
"""
import calendar
import os
import pickle
import time
from pathlib import Path

import pandas as pd

# Excel modunda oracledb kurulu olmasa da rapor üretilebilsin diye içe
# aktarma hatası burada değil, bağlanırken bildirilir.
try:
    import oracledb
except ImportError:  # pragma: no cover
    oracledb = None


def db_connection():
    if oracledb is None:
        raise ConnectionError("oracledb kurulu değil: pip install oracledb")
    eksik = [a for a in ("ORACLE_USER", "ORACLE_PASSWORD", "ORACLE_DSN") if not os.environ.get(a)]
    if eksik:
        raise ConnectionError(f"Oracle bağlantı bilgisi eksik, ortam değişkenlerini tanımlayın: {eksik}")
    try:
        return oracledb.connect(user=os.environ["ORACLE_USER"], password=os.environ["ORACLE_PASSWORD"],
                                dsn=os.environ["ORACLE_DSN"])
    except Exception as e:
        raise ConnectionError(f"Failed to establish database connection: {e}")


# arraysize: python-oracledb varsayılan olarak bir seferde 100 satır ister;
# büyük tabloda binlerce ağ gidiş-dönüşü demektir. Büyütmek en büyük kazanç.
ORACLE_ARRAYSIZE = 5000        # sunucudan bir seferde istenen satır sayısı
ORACLE_PARCA_BOYUTU = 100_000  # belleğe alınan parça büyüklüğü

# Aynı çalıştırma içinde tekrarlanan sorgular buraya yazılır (iki kaynak aynı
# tabloyu okuyorsa tablo bir kez indirilir).
_ORACLE_ONBELLEK: dict = {}


def fetch_data_hizli(query, onbellege_al: bool = True):
    """fetch_data ile AYNI sözleşme (DataFrame döndürür); sadece daha hızlı."""
    anahtar = " ".join(str(query).split())
    if onbellege_al and anahtar in _ORACLE_ONBELLEK:
        print("[BİLGİ] Bu sorgu bu çalıştırmada ZATEN çekilmişti, önbellekten kullanılıyor.")
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
            print(f"[BİLGİ]   ... {toplam:,} satır".replace(",", ".") + f" ({time.time() - basladi:.0f} sn)")
        cursor.close()

        df = pd.concat(parcalar, ignore_index=True) if parcalar else pd.DataFrame(columns=columns)
        print(f"[BİLGİ] {len(df):,} satır x {len(df.columns)} kolon".replace(",", ".")
              + f", {time.time() - basladi:.0f} saniyede çekildi (arraysize={ORACLE_ARRAYSIZE}).")
        if onbellege_al:
            _ORACLE_ONBELLEK[anahtar] = df
        return df
    except Exception as e:
        raise RuntimeError(f"Query execution failed: {e}")
    finally:
        if conn:
            conn.close()


# Excel'den okuma: `VERI_KAYNAGI = "excel"` iken veri Excel dosyasından okunur (Oracle ile sayıları karşılaştırmak için birebir). Büyük dosyalar ilk okumadan sonra yanına yazılan `.cache.pkl` önbelleğinden saniyeler içinde okunur; dosya değişince önbellek kendiliğinden yenilenir.
def _read_excel_fast(path: Path) -> pd.DataFrame:
    """Büyük Excel dosyalarını okunabilir sürede okur: önbellek → calamine → openpyxl."""
    cache_path = path.with_suffix(path.suffix + ".cache.pkl")
    try:
        stat = path.stat()
        signature = (str(path.resolve()), stat.st_size, int(stat.st_mtime))
    except OSError:
        signature = None

    if signature is not None and cache_path.exists():
        try:
            with open(cache_path, "rb") as fh:
                cached_signature, cached_df = pickle.load(fh)
            if cached_signature == signature:
                print(f"[BİLGİ] {path.name}: önbellekten okundu ({len(cached_df)} satır).")
                return cached_df
            print(f"[BİLGİ] {path.name}: veri dosyası değişmiş, önbellek yenileniyor.")
        except Exception as exc:  # noqa: BLE001
            print(f"[BİLGİ] {path.name}: önbellek okunamadı ({exc}), dosya yeniden okunuyor.")

    size_mb = path.stat().st_size / (1024 * 1024) if signature is not None else 0
    print(f"[BİLGİ] {path.name} okunuyor ({size_mb:.0f} MB)... Büyük dosyalarda bu adım dakikalar sürebilir.")
    started = time.time()
    df = None
    try:
        import python_calamine  # noqa: F401
    except ImportError:
        pass
    else:
        try:
            df = pd.read_excel(path, engine="calamine")
        except Exception as exc:  # noqa: BLE001
            print(f"[BİLGİ] {path.name}: 'calamine' motoru kullanılamadı ({exc}), openpyxl'e dönülüyor.")
            df = None
    if df is None:
        df = pd.read_excel(path)
    print(f"[BİLGİ] {path.name}: {len(df)} satır x {len(df.columns)} kolon, {time.time() - started:.0f} saniyede okundu.")

    if signature is not None:
        try:
            with open(cache_path, "wb") as fh:
                pickle.dump((signature, df), fh, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception as exc:  # noqa: BLE001
            print(f"[BİLGİ] {path.name}: önbellek yazılamadı ({exc}); rapor yine üretilir.")
    return df


def _ay_sonu(donem: int) -> str:
    y, m = divmod(int(donem), 100)
    return f"TO_DATE('{calendar.monthrange(y, m)[1]:02d}/{m:02d}/{y}','DD/MM/YYYY')"


def _sorguyu_doldur(sorgu: str, aylar) -> str:
    return (str(sorgu).replace("{DONEMLER}", ", ".join(str(a) for a in aylar))
            .replace("{AY_SONLARI}", ", ".join(_ay_sonu(a) for a in aylar)))


def _sorgu_dolu_mu(sorgu: str) -> bool:
    govde = "\n".join(s for s in str(sorgu).splitlines() if not s.strip().startswith("--"))
    return bool(govde.strip()) and "DOLDURULACAK" not in str(sorgu)


_YER_TUTUCULAR = ("{DONEMLER}", "{AY_SONLARI}")


def kaynak_oku(ayar: dict, mod: str = "excel", aylar=None) -> pd.DataFrame:
    """Bir kaynağın verisini DataFrame olarak döndürür.

    ayar: {"ad": ..., "oracle_sorgu": ..., "excel_dosyasi": ...}
    mod : "oracle" ya da "excel"
    aylar: sorguda {DONEMLER} / {AY_SONLARI} varsa yerine konacak dönemler
           (ör. WHERE DONEM IN ({DONEMLER})).
    """
    mod = mod.strip().lower()
    if mod == "oracle":
        sorgu = ayar["oracle_sorgu"]
        if not _sorgu_dolu_mu(sorgu):
            raise RuntimeError(f"'{ayar['ad']}' için Oracle sorgusu doldurulmamış.")
        if any(y in str(sorgu) for y in _YER_TUTUCULAR):
            if not aylar:
                raise RuntimeError(f"'{ayar['ad']}' sorgusu raporun aylarını istiyor ama aylar verilmedi.")
            sorgu = _sorguyu_doldur(sorgu, aylar)
        return fetch_data_hizli(sorgu)
    if mod == "excel":
        yol = Path(ayar["excel_dosyasi"])
        if not yol.exists():
            raise FileNotFoundError(f"'{ayar['ad']}' veri dosyası bulunamadı: {yol.resolve()}")
        return _read_excel_fast(yol)
    raise ValueError(f"mod 'oracle' ya da 'excel' olmalı, '{mod}' verildi.")
