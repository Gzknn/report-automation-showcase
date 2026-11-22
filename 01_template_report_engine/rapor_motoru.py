# %% İçe aktarmalar | Motorun kullandığı standart kütüphaneler. Motor yalnızca `openpyxl` ve (veriyi hazırlarken) `pandas` ister.
"""Şablon tabanlı rapor motoru.

Boş bir Excel şablonu + okunur bir kural sözlüğü (`RAPOR`) + veri alır,
şablondaki tabloları (panelleri) başlıklarından keşfeder ve doldurur.
Kodda satır/sütun numarası gömülü değildir; şablon değişince kod değişmez.

Motor, tarayıcıda çalışan bir önizleme motorunun (JavaScript) Python
karşılığı olarak yazıldı: ikisi AYNI kural sözlüğünü AYNI anlamla yorumlamak
zorunda olduğu için JavaScript'in sayı/metin davranışları (String(),
Number(), Math.round) aşağıdaki yardımcılarla taklit edilir.
"""
import math
import re
from decimal import Decimal

# %% JavaScript ile aynı sayı ve metin davranışı | Veri Oracle'dan `Decimal`, Excel'den `1.0` gibi ondalıklı gelebilir; önizlemedeki JavaScript ise `1` görür. Karşılaştırmalar bu yardımcılarla aynı metne indirilir.


def _js_ondalik(x):
    """JavaScript'in Number#toString'i: 1e21'e kadar üstel yazmaz, 1e-7 gibi yazar."""
    d = Decimal(repr(x))
    isaret, rakamlar, us = d.as_tuple()
    rakamlar = "".join(map(str, rakamlar)).rstrip("0") or "0"
    k = len(rakamlar)
    n = us + len("".join(map(str, d.as_tuple()[1])))  # ondalık noktanın konumu
    on = "-" if isaret else ""
    if k <= n <= 21:
        return on + rakamlar + "0" * (n - k)
    if 0 < n <= 21:
        return on + rakamlar[:n] + "." + rakamlar[n:]
    if -6 < n <= 0:
        return on + "0." + "0" * (-n) + rakamlar
    e = n - 1
    govde = rakamlar if k == 1 else rakamlar[0] + "." + rakamlar[1:]
    return on + govde + "e" + ("+" if e > 0 else "-") + str(abs(e))


def js_metin(v):
    """JavaScript String(v)."""
    if v is None:
        return "null"
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, Decimal):
        v = int(v) if v == v.to_integral_value() else float(v)
        return js_metin(v)
    if isinstance(v, float):
        if v != v:
            return "NaN"
        if math.isinf(v):
            return "Infinity" if v > 0 else "-Infinity"
        if v.is_integer() and abs(v) < 1e21:
            return str(int(v))
        return _js_ondalik(v)
    return str(v)


def js_sayi(v):
    """JavaScript Number(v); sayı değilse NaN."""
    if v is None:
        return 0
    if isinstance(v, bool):
        return 1 if v else 0
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, Decimal):
        return int(v) if v == v.to_integral_value() else float(v)
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return 0
        if "_" in s:
            return math.nan
        try:
            return float(s)
        except ValueError:
            return math.nan
    return math.nan


# Aynı değer milyonlarca kez metne çevrilir; sonuç önbellekte tutulur. Anahtar
# türü de içerir: 1, 1.0 ve True Python'da aynı anahtardır ama JS'te "1",
# "1" ve "true" olur.
_METIN_ONBELLEK = {}


def metin(v):
    try:
        return _METIN_ONBELLEK[(type(v), v)]
    except KeyError:
        pass
    except TypeError:  # hash'lenemeyen değer
        return "" if v is None else js_metin(v).strip()
    m = "" if v is None or (isinstance(v, float) and v != v) else js_metin(v).strip()
    if len(_METIN_ONBELLEK) < 1_000_000:
        _METIN_ONBELLEK[(type(v), v)] = m
    return m


def bos_mu(v):
    return v is None or (isinstance(v, float) and v != v) or metin(v) == ""


def _sonlu(x):
    return isinstance(x, (int, float)) and not (isinstance(x, float) and (x != x or math.isinf(x)))


# %% Dönem penceresi | Rapor verideki en güncel aydan geriye, şablondaki dönem sayısı kadar ay gösterir. Yıl taşması (202512 → 202601) `ay_ekle` ile yönetilir.


def donem_mi(v):
    if isinstance(v, bool) or not isinstance(v, int):
        return False
    s = str(v)
    if not re.fullmatch(r"\d{6}", s):
        return False
    y, m = int(s[:4]), int(s[4:])
    return 2000 <= y <= 2099 and 1 <= m <= 12


def ay_ekle(donem, fark):
    y, m = donem // 100, donem % 100
    t = y * 12 + (m - 1) + fark
    return (t // 12) * 100 + (t % 12) + 1


def pencere(hedef, n):
    return [ay_ekle(hedef, -(n - 1) + i) for i in range(n)]


# %% Şablon ve panel keşfi | Şablondaki tablolar (paneller) başlıklarından ve dönem sütunlarından bulunur; kodda satır numarası yoktur. Okuma, yazılmamış (bozulmamış) bir şablon kopyasından yapılır: önizlemedeki motor da hep şablonun kendisine bakar.


class SayfaModeli:
    """openpyxl sayfasının motorun okuduğu hâli: değer, sayı biçimi, birleşik aralıklar.

    Değerler araclar/sablon_json.py'deki gibi çevrilir (tarih → metin), ki
    önizlemenin gördüğü şablonla birebir aynı olsun.
    """

    def __init__(self, ws):
        self.ad = ws.title
        self.maxRow = ws.max_row
        self.maxCol = ws.max_column
        self.birlesik = [[m.min_row, m.min_col, m.max_row, m.max_col] for m in ws.merged_cells.ranges]
        self._deger, self._bicim = {}, {}
        for satir in ws.iter_rows(min_row=1, max_row=self.maxRow, max_col=self.maxCol):
            for h in satir:
                v = h.value
                if v is not None:
                    self._deger[(h.row, h.column)] = v if isinstance(v, (int, float, str)) else str(v)
                if h.number_format and h.number_format != "General":
                    self._bicim[(h.row, h.column)] = h.number_format

    def deger(self, r, c):
        return self._deger.get((r, c))

    def bicim(self, r, c):
        return self._bicim.get((r, c), "General")


SAHTE_BASLIK = re.compile(r"^\d{6}\s*(adet|oran)", re.IGNORECASE)
ALT_BASLIK = re.compile(r"^\s*(adet|adetler|oran|oranlar)\s*$", re.IGNORECASE)
GRAND_TOTAL = re.compile(r"grand\s*total", re.IGNORECASE)


def _baslik_uygun(v):
    if v is None or donem_mi(v):
        return False
    m = js_metin(v).strip()
    return m != "" and not SAHTE_BASLIK.match(m)


def _basliklari_bul(sayfa, r):
    t = []
    for b in sayfa.birlesik:
        if b[0] == r and b[2] == r + 1 and b[1] == b[3]:
            v = sayfa.deger(r, b[1])
            if _baslik_uygun(v):
                t.append([b[1], js_metin(v), False])
    if t:
        return sorted(t, key=lambda x: x[0])
    for c in range(1, sayfa.maxCol + 1):
        v = sayfa.deger(r, c)
        if isinstance(v, str) and _baslik_uygun(v) and donem_mi(sayfa.deger(r, c + 1)):
            t.append([c, v, False])
    if t:
        return t
    for c in range(1, sayfa.maxCol + 1):
        if sayfa.deger(r, c) is not None:
            continue
        if not donem_mi(sayfa.deger(r, c + 1)):
            continue
        if c > 1 and donem_mi(sayfa.deger(r, c - 1)):
            continue
        alt = sayfa.deger(r + 1, c)
        if not isinstance(alt, str) or not _baslik_uygun(alt):
            continue
        t.append([c, alt, True])
    return t


def sutun_harfi(c):
    s = ""
    while c > 0:
        m = (c - 1) % 26
        s = chr(65 + m) + s
        c = (c - 1) // 26
    return s


def panelleri_kesfet(sayfa):
    paneller = []
    for r in range(1, sayfa.maxRow + 1):
        basliklar = _basliklari_bul(sayfa, r)
        if not basliklar:
            continue
        satir_paneller = [{"baslik": b[1].strip(), "headerRow": r, "labelCol": b[0], "altta": b[2],
                           "periodCols": [], "labelRows": []} for b in basliklar]
        for c in range(1, sayfa.maxCol + 1):
            if not donem_mi(sayfa.deger(r, c)):
                continue
            sahip = None
            for p in satir_paneller:
                if p["labelCol"] <= c:
                    sahip = p
            (sahip or satir_paneller[0])["periodCols"].append(c)
        for i, p in enumerate(satir_paneller):
            p["rightBoundary"] = satir_paneller[i + 1]["labelCol"] - 1 if i + 1 < len(satir_paneller) else sayfa.maxCol
        for p in satir_paneller:
            if not p["periodCols"]:
                continue
            # Başlığın altındaki satır bir alt başlıksa (boş, "Adet"/"Oran" ya
            # da dönem sütununda "Adet") veri bir satır aşağıdan başlar.
            if p["altta"]:
                bas = r + 2
            else:
                v = sayfa.deger(r + 1, p["labelCol"])
                dv = sayfa.deger(r + 1, p["periodCols"][0])
                bas = r + 2 if (v is None or ALT_BASLIK.search(js_metin(v)) or (isinstance(dv, str) and ALT_BASLIK.search(dv))) else r + 1
            s = bas
            while s <= sayfa.maxRow and sayfa.deger(s, p["labelCol"]) is not None:
                p["labelRows"].append(s)
                s += 1
            p["gtRow"] = None
            for s in p["labelRows"]:
                if GRAND_TOTAL.search(js_metin(sayfa.deger(s, p["labelCol"]))):
                    p["gtRow"] = s
                    break
            paneller.append(p)
    # Anahtar: başlık; tekrar ediyorsa okuma sırasıyla "Panel (2)".
    sayac = {}
    for p in paneller:
        sayac[p["baslik"]] = sayac.get(p["baslik"], 0) + 1
        p["anahtar"] = p["baslik"] if sayac[p["baslik"]] == 1 else f"{p['baslik']} ({sayac[p['baslik']]})"
        p["adres"] = sutun_harfi(p["labelCol"]) + str(p["headerRow"])
    return paneller


def grup_sinirlari(p):
    """Son grup diğerlerinin genişliği kadardır (sağdaki TOTAL grubuna taşmasın)."""
    pc = p["periodCols"]
    gen = pc[1] - pc[0] if len(pc) > 1 else None
    return [[c, pc[i + 1] - 1 if i + 1 < len(pc) else (min(c + gen - 1, p["rightBoundary"]) if gen else p["rightBoundary"])]
            for i, c in enumerate(pc)]


def toplam_grubu(sayfa, p):
    """Dönem sütunlarının SAĞINDAKİ toplam grubu ("TOTAL"): son dönem grubundan
    sonra başlık satırında yazısı olan ilk sütundan sağ sınıra kadar."""
    g = grup_sinirlari(p)
    if not g:
        return None
    for c in range(g[-1][1] + 1, p["rightBoundary"] + 1):
        v = sayfa.deger(p["headerRow"], c)
        if v is not None and str(v).strip() != "":
            return [c, p["rightBoundary"]]
    return None


def sutunlari_siniflandir(sayfa, ornek_satir, bas, son, genel_adet=False):
    """genel_adet: biçimi General olan sütunlar da adet (yalnızca yüzde biçimliler oran)."""
    sonuc = []
    for c in range(bas, son + 1):
        f = sayfa.bicim(ornek_satir, c)
        sonuc.append([c, "oran" if "%" in f else ("adet" if (f != "General" or genel_adet) else None)])
    return sonuc


def adet_oran_eslestir(siniflar):
    bekleyen, ciftler = [], []
    for c, tur in siniflar:
        if tur == "adet":
            bekleyen.append(c)
        elif tur == "oran" and bekleyen:
            ciftler.append([c, bekleyen.pop(0)])
    return ciftler


def sablon_aylari(sayfa, paneller):
    p = next((x for x in paneller if x["periodCols"]), None)
    if not p:
        return []
    return sorted(sayfa.deger(p["headerRow"], c) for c in p["periodCols"])


def ozet_govde(sayfa, p):
    c = p["periodCols"][0]
    return [r for r in p["labelRows"] if sayfa.bicim(r, c) != "General" and "%" not in sayfa.bicim(r, c)]


# %% Sayım ve filtre | Tek sayım noktası `topla`: adet kolonu varsa onun toplamı, yoksa kayıt sayısı. Filtre koşulları metin olarak karşılaştırılır (Excel'den `1`, Oracle'dan `'1'` gelse de aynı).


def _iceriyor(liste, v):
    return any(metin(x) == v for x in liste)


def kosul_uyar(satir, k):
    ist = k.get("istisna")
    if ist and _iceriyor(ist["degerler"], metin(satir.get(ist["kolon"]))):
        return True
    var_mi = _iceriyor(k["degerler"], metin(satir.get(k["kolon"])))
    return not var_mi if k.get("op") == "not_in" else var_mi


def suz(satirlar, kosullar):
    if not kosullar:
        return satirlar
    # kosul_uyar'ın aynısı; değer kümeleri satır başına değil bir kez kurulur.
    hazir = []
    for k in kosullar:
        ist = k.get("istisna")
        hazir.append((k["kolon"], {metin(x) for x in k["degerler"]}, k.get("op") == "not_in",
                      ist["kolon"] if ist else None, {metin(x) for x in ist["degerler"]} if ist else None))

    def uyar(s):
        for kolon, degerler, degil, ist_kolon, ist_degerler in hazir:
            if ist_kolon is not None and metin(s.get(ist_kolon)) in ist_degerler:
                continue
            if (metin(s.get(kolon)) in degerler) == degil:
                return False
        return True
    return [s for s in satirlar if uyar(s)]


def agirlik_fn(kaynak):
    ak = kaynak.get("adetKolonu")
    if not ak:
        return lambda s: 1

    def f(s):
        x = js_sayi(s.get(ak))
        return x if _sonlu(x) else 0
    return f


def topla(satirlar, agirlik):
    t = 0
    for s in satirlar:
        t += agirlik(s)
    return t


def yuvarla(x):
    """Önizleme motorundaki yuvarla: yarımlar çifte (Python round() gibi)."""
    a = math.floor(x)
    f = x - a
    if abs(f - 0.5) < 1e-9:
        return a if a % 2 == 0 else a + 1
    return math.floor(x + 0.5)


def normallestir(v):
    s = re.sub(r"[İIı]", "i", js_metin(v)).lower()
    return re.sub(r"[-_\s]+", " ", s).strip()


def kosullu_alan(satirlar, g):
    """Hesaplanan alan (CASE WHEN): kurallar sırayla, ilk tutanın adı yazılır.

    Kural = seçenekler (VEYA), seçenek = koşullar (VE). Önce yazılan kural
    önceliklidir (ör. Onay, Ret'ten önce).
    """
    kurallar = g.get("kurallar") or []
    for s in satirlar:
        v = None
        for kural in kurallar:
            if any(all(kosul_uyar(s, k) for k in secenek) for secenek in kural.get("secenekler") or []):
                v = kural["ad"]
                break
        s[g["ad"]] = g["varsayilan"] if v is None and "varsayilan" in g else v


def gruplari_uygula(satirlar, gruplar):
    """Excel pivot'taki "Grupla": ham değerleri tek ada toplayan yeni kolon."""
    for g in gruplar or []:
        if g.get("tip") == "kosullu":
            kosullu_alan(satirlar, g)
            continue
        harita, norm_harita = {}, {}
        for grup in g["gruplar"]:
            for d in grup["degerler"]:
                harita[metin(d)] = grup["ad"]
                if g.get("normalize"):
                    norm_harita[normallestir(d)] = grup["ad"]
            if g.get("normalize"):
                norm_harita[normallestir(grup["ad"])] = grup["ad"]
        for s in satirlar:
            ham = s.get(g["kolon"])
            if bos_mu(ham):
                s[g["ad"]] = ham
                continue
            h = harita.get(metin(ham))
            if h is None and g.get("normalize"):
                h = norm_harita.get(normallestir(ham))
            s[g["ad"]] = js_metin(ham) if h is None else h
    return satirlar


def kosul_metni(k):
    s = k["kolon"] + (" ∉ " if k.get("op") == "not_in" else " ∈ ") + "{" + ", ".join(map(str, k["degerler"])) + "}"
    if k.get("istisna"):
        s += f" ({k['istisna']['kolon']} ∈ {{{', '.join(map(str, k['istisna']['degerler']))}}} hariç)"
    return s


# %% Satır ağacı | Tablonun satırları kural sözlüğünde alt alta konan alanlardan kurulur; açık değerlerin altına sonraki alan gelir. `agactan_bloklar` bu ağacı şablon satırlarını paylaşan parçalara çevirir, `girdileri_belirle` her parçada hangi değerin gösterileceğini (sabit liste, en çok N + Diğer, A→Z) belirler.


def alt_satir_sayisi(alt):
    if not alt or alt.get("mod") == "sablon":
        return 0
    if alt.get("satirSayisi"):
        return alt["satirSayisi"]
    return len(alt.get("sabit") or []) if alt.get("mod") == "sabit" else 0


def blok_ihtiyaci(b):
    if b.get("satirSayisi"):
        return b["satirSayisi"]
    if b.get("alt") and b["alt"].get("mod") == "sablon":
        return None
    if b.get("mod") == "sabit":
        return len(b.get("sabit") or []) * (1 + alt_satir_sayisi(b.get("alt")))
    return None


def satirlari_dagit(bloklar, govde):
    sabit_toplam, otomatik = 0, []
    for i, b in enumerate(bloklar):
        n = blok_ihtiyaci(b)
        if n is None:
            otomatik.append(i)
        else:
            sabit_toplam += n
    kalan = max(0, len(govde) - sabit_toplam)
    pay = kalan // len(otomatik) if otomatik else 0
    artik = kalan - pay * len(otomatik) if otomatik else 0
    bas, dagitim = 0, []
    for i, b in enumerate(bloklar):
        n = blok_ihtiyaci(b)
        if n is None:
            n = pay + (artik if otomatik[-1] == i else 0)
        dagitim.append(govde[bas:bas + n])
        bas += n
    uyari = None
    if sabit_toplam > len(govde):
        uyari = f"Satırlar {sabit_toplam} şablon satırı istiyor ama tabloda {len(govde)} satır var; fazlası yazılmadı."
    elif artik and len(otomatik) > 1:
        uyari = f"Kalan {kalan} satır, açılmış {len(otomatik)} gruba eşit bölünemedi; artan {artik} satır sonuncuya verildi."
    return {"dagitim": dagitim, "uyari": uyari, "kullanilmayan": govde[bas:]}


def girdileri_belirle(tanim, havuz_pencere, son_ay, donem_kolonu, agirlik, yer, aylar=None):
    kolon = tanim["satir"]
    haric = [metin(x) for x in tanim.get("haric") or []]
    havuz = [s for s in havuz_pencere if metin(s.get(kolon)) not in haric]
    diger = tanim.get("diger") or {}
    diger_var = bool(diger.get("aktif")) and tanim.get("mod") != "sabit"
    yaprak_yer = max(0, yer - 1) if diger_var else yer

    if tanim.get("mod") == "sabit":
        aday_sayisi = len(tanim.get("sabit") or [])
        girdiler = [{"etiket": x.get("etiket") or x["degerler"][0], "degerler": [metin(d) for d in x["degerler"]]}
                    for x in (tanim.get("sabit") or [])[:yaprak_yer]]
    else:
        sayim = {}
        # bos_etiket: boş değerler "BOŞ" gibi tek satırda gösterilir; kademeli
        # sıralama ve bos_etiket değerleri kırpılmış metinle anahtarlar.
        bos_etiket = tanim.get("bosEtiket") or ""
        kademeli = tanim.get("mod") == "ilk_n" and tanim.get("siralama") == "kademeli"
        anahtar_of = metin if (bos_etiket or kademeli) else js_metin

        def ad_of(k):
            return bos_etiket if k == "" and bos_etiket else k
        siralama_havuzu = havuz
        if tanim.get("mod") == "ilk_n" and tanim.get("siralama") == "son_ay":
            siralama_havuzu = [s for s in havuz if js_sayi(s.get(donem_kolonu)) == son_ay]
        aylik = {}
        for s in siralama_havuzu:
            v = s.get(kolon)
            if bos_mu(v) and not bos_etiket:
                continue
            k = "" if bos_mu(v) else anahtar_of(v)
            sayim[k] = sayim.get(k, 0) + agirlik(s)
            if kademeli:
                ay = js_sayi(s.get(donem_kolonu))
                m = aylik.setdefault(ay, {})
                m[k] = m.get(k, 0) + agirlik(s)
        anahtarlar = list(sayim)
        aday_sayisi = len(anahtarlar)
        if tanim.get("mod") == "alfabetik":
            anahtarlar.sort(key=ad_of)
        elif kademeli:
            # Son ayın adedi; eşitse bir önceki ay, ...; hiçbirinde ayrışmazsa ad.
            ters = list(reversed(aylar or [son_ay]))
            anahtarlar.sort(key=lambda a: tuple(-aylik.get(ay, {}).get(a, 0) for ay in ters) + (ad_of(a),))
        else:
            anahtarlar.sort(key=lambda a: (-sayim[a], ad_of(a)))
        secilen = anahtarlar[:yaprak_yer]
        # Her zaman gösterilecek değerler: veride varsa ilk N'e girmeseler
        # bile sabitlenmemiş en küçük değerin yerine geçer.
        sabitle = [str(v) for v in (tanim.get("sabitle") or [] if tanim.get("mod") == "ilk_n" else []) if str(v) in sayim]
        for v in sabitle:
            if v in secilen:
                continue
            for i in range(len(secilen) - 1, -1, -1):
                if secilen[i] not in sabitle:
                    del secilen[i]
                    secilen.append(v)
                    break
        girdiler = [{"etiket": ad_of(k), "degerler": [metin(k)]} for k in anahtarlar if k in secilen]
    while len(girdiler) < yaprak_yer:
        girdiler.append({"bos": True})
    if diger_var and yer > 0:
        girdiler.append({"etiket": diger.get("etiket") or "Diğer", "diger": True})
    return {"girdiler": girdiler, "haric": haric, "adaySayisi": aday_sayisi, "yaprakYer": yaprak_yer}


def dugum_anahtari(yol):
    return "›".join(yol)


def agactan_bloklar(agac, gruplar=None):
    sev = (agac or {}).get("seviyeler") or []
    acik = (agac or {}).get("acik") or {}
    bloklar, hatalar, notlar, kullanilan = [], [], [], set()

    # Excel pivot'taki gibi: gruplanmış alanın değeri açılıp altına grubun
    # kaynak kolonu konunca altta yalnızca o değerin üyeleri listelenir
    # (önizleme motorundaki agactanBloklar ile aynı).
    grup_of = {g["ad"]: g for g in gruplar or [] if g.get("tip") != "kosullu" and g.get("gruplar") is not None}

    def kapsam_bul(L, sonraki, girdi):
        g = grup_of.get(L.get("satir"))
        if not g or not sonraki or sonraki.get("mod") != "sabit" or g.get("kolon") != sonraki.get("satir"):
            return None
        norm = normallestir if g.get("normalize") else metin
        uye = set()
        for v in girdi["degerler"]:
            gr = next((x for x in g["gruplar"] if metin(x["ad"]) == metin(v)), None)
            uye.update(norm(x) for x in (gr["degerler"] if gr else [v]))
        return lambda g2: any(norm(x) in uye for x in g2["degerler"])

    def payda_cevir(pd, derinlik, alt_mi):
        pd = pd or ("ust" if derinlik else "grand_total")
        if pd == "kardes":
            return "blok"
        if pd == "ust":
            return "ust" if alt_mi else ("blok_ham" if derinlik else "grand_total")
        return pd

    def parca(L, d, kosullar, ad):
        b = {"ad": ad, "satir": L.get("satir"), "mod": L.get("mod") or "sabit", "filtre": list(kosullar),
             "derinlik": d, "payda": payda_cevir(L.get("payda"), d, False)}
        for k in ("siralama", "sabitle", "diger", "haric", "bosEtiket"):
            if k in L:
                b[k] = L[k]
        return b

    def yuru(d, kosullar, yol, ust_ad, kapsam=None):
        L = sev[d]
        sonraki = sev[d + 1] if d + 1 < len(sev) else None
        kullanilan.add(d)
        ad = (ust_ad + " › " if ust_ad else "") + (L.get("satir") or "?")
        if not L.get("satir"):
            hatalar.append(f"{d + 1}. seviyeye alan seçilmedi.")
            return
        if L.get("mod") == "sablon":
            hatalar.append(f"'{L['satir']}' şablondan okunan gruplar ancak değişken satırlı bir seviyenin altında (her birinin altında) kullanılabilir.")
            return
        if L.get("mod") != "sabit":
            b = parca(L, d, kosullar, ad)
            if L.get("satirSayisi"):
                b["satirSayisi"] = L["satirSayisi"]
            if sonraki and L.get("altAcik"):
                kullanilan.add(d + 1)
                if not sonraki.get("satir"):
                    hatalar.append(f"{d + 2}. seviyeye alan seçilmedi.")
                elif sonraki.get("mod") not in ("sabit", "sablon") and not sonraki.get("satirSayisi"):
                    hatalar.append(f"'{L['satir']}' değerleri veriden geldiği için altındaki '{sonraki['satir']}' ya sabit liste, ya şablondan okunan grup olmalı ya da satır sayısı verilmeli.")
                else:
                    b["alt"] = {"satir": sonraki["satir"], "mod": sonraki.get("mod") or "sabit",
                                "payda": payda_cevir(sonraki.get("payda"), d + 1, True)}
                    for k in ("sabit", "siralama", "sabitle", "diger", "haric", "satirSayisi", "bosEtiket"):
                        if k in sonraki:
                            b["alt"][k] = sonraki[k]
            bloklar.append(b)
            return

        bekleyen = []

        def bosalt():
            if not bekleyen:
                return
            b = parca(L, d, kosullar, ad)
            b["sabit"] = list(bekleyen)
            bloklar.append(b)
            bekleyen.clear()

        for girdi in [x for x in L.get("sabit") or [] if not kapsam or kapsam(x)]:
            y = yol + ["|".join(metin(x) for x in girdi["degerler"])]
            bekleyen.append(girdi)
            if sonraki and acik.get(dugum_anahtari(y)):
                bosalt()
                yuru(d + 1, kosullar + [{"kolon": L["satir"], "op": "in", "degerler": [metin(x) for x in girdi["degerler"]]}],
                     y, girdi.get("etiket") or girdi["degerler"][0], kapsam_bul(L, sonraki, girdi))
        bosalt()

    if sev:
        yuru(0, [], [], "")
    for i, L in enumerate(sev):
        if i not in kullanilan:
            notlar.append(f"{i + 1}. seviye ('{L.get('satir') or '?'}') hiçbir değerin altında açık değil; tabloda görünmüyor.")
    return {"bloklar": bloklar, "hata": " ".join(hatalar) if hatalar else None, "notlar": notlar}


# %% Sayfa doldurma | Her tablo kendi kuralıyla doldurulur; kuralı olmayan tablo atlanır ve sebebi yazılır. Sonuç `yazilanlar` sözlüğüdür: {(satır, sütun): değer}. Hücrelere en sonda, tek seferde yazılır.


def sayfa_doldur(sayfa, paneller, sayfa_kurali, rapor, veri, aylar):
    yazilanlar = {}
    # Paketlenen tablolarda taşınan satırların biçimi: (r, c) -> şablondaki (r0, c0)
    stiller = {}
    mesajlar = []
    sonuc = {"dolu": [], "kuralsiz": [], "hata": [], "yazilamadi": []}
    kurallar = (sayfa_kurali or {}).get("paneller") or {}

    def yaz(r, c, v):
        yazilanlar[(r, c)] = v

    def mesaj(seviye, panel, m):
        mesajlar.append((seviye, panel, m))

    def basliklari_yaz(p):
        for i, g in enumerate(grup_sinirlari(p)):
            if i >= len(aylar):
                continue
            ay = aylar[i]
            eski = sayfa.deger(p["headerRow"], g[0])
            yaz(p["headerRow"], g[0], ay)
            if donem_mi(eski):
                for c in range(g[0] + 1, g[1] + 1):
                    v = sayfa.deger(p["headerRow"], c)
                    if isinstance(v, str) and str(eski) in v:
                        yaz(p["headerRow"], c, v.replace(str(eski), str(ay)))

    ay_onbellek = {}

    def ayin_satirlari(kaynak_id, ay):
        k = rapor["kaynaklar"].get(kaynak_id)
        if not k:
            return []
        anahtar = (kaynak_id, ay)
        if anahtar not in ay_onbellek:
            dk = k["donemKolonu"]
            ay_onbellek[anahtar] = [s for s in veri.get(kaynak_id) or [] if js_sayi(s.get(dk)) == ay]
        return ay_onbellek[anahtar]

    def ozel_satirlari(kural):
        if kural.get("ozelSatirlar"):
            return {int(k): v for k, v in kural["ozelSatirlar"].items()}
        return {i: s for i, s in enumerate(kural.get("satirlar") or [])}

    def ozel_uygula(p, kural, govde, sinirlar):
        ozel = ozel_satirlari(kural)
        sira = sorted(ozel)
        fazla = [i for i in sira if i >= len(govde)]
        if fazla:
            mesaj("uyari", p["anahtar"], f"{len(fazla)} özel satır şablondaki veri satırlarının dışında kaldı; yok sayıldı.")
        sira = [i for i in sira if i < len(govde)]
        if not sira:
            return

        def kitle(kaynak_id, filtre, ay):
            satirlar = ayin_satirlari(kaynak_id, ay)
            if kaynak_id == kural.get("kaynak"):
                satirlar = suz(satirlar, kural.get("filtre"))
            return suz(satirlar, filtre)

        def kolon_toplami(kit, kolon):
            t = 0
            for x in kit:
                n = js_sayi(x.get(kolon))
                t += n if n == n else 0  # JS: Number(x) || 0
            return t

        def olc(s, ay):
            kaynak_id = s.get("kaynak") or kural.get("kaynak")
            k = rapor["kaynaklar"].get(kaynak_id)
            if not k:
                raise ValueError(f"'{kaynak_id}' veri kaynağı tanımlı değil.")
            kit = kitle(kaynak_id, s.get("filtre"), ay)
            o = s.get("olcu") or {"tip": "adet"}
            bolen = o.get("bolen") or 1
            if o.get("tip") == "toplam":
                return yuvarla(kolon_toplami(kit, o["kolon"]) / bolen)
            if o.get("tip") == "ortalama":
                pay = kolon_toplami(kit, o["kolon"])
                payda = topla(kit, agirlik_fn(k))
                return yuvarla(pay / payda / bolen) if payda else 0
            return topla(kit, agirlik_fn(k))

        # Gizlenen satırda ağacın yazdıkları da silinir: şablon olduğu gibi kalır.
        for i in sira:
            if ozel[i].get("atla"):
                for anahtar in [a for a in yazilanlar if a[0] == govde[i]]:
                    del yazilanlar[anahtar]
        for i in sira:
            if not ozel[i].get("atla") and ozel[i].get("etiket"):
                yaz(govde[i], p["labelCol"], ozel[i]["etiket"])

        for ai, ay in enumerate(aylar):
            sinir = sinirlar[ai]
            ornek = next((r for r in govde if any("%" in sayfa.bicim(r, c) for c in range(sinir[0], sinir[1] + 1))), govde[0])
            siniflar = sutunlari_siniflandir(sayfa, ornek, sinir[0], sinir[1], kural.get("genelAdet"))
            ciftler = adet_oran_eslestir(siniflar)
            adet_sutunlari = [c for c, t in siniflar if t == "adet"]
            degerler = {i: olc(ozel[i], ay) for i in sira if not ozel[i].get("atla")}

            def satir_degeri(j):
                if j in degerler:
                    return degerler[j]
                if 0 <= j < len(govde) and adet_sutunlari:
                    v = yazilanlar.get((govde[j], adet_sutunlari[0]))
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        return v
                return 0

            for i in sira:
                s, r = ozel[i], govde[i]
                if s.get("atla"):
                    continue
                v = degerler[i]
                for c in adet_sutunlari:
                    yaz(r, c, v)
                pd = s.get("payda")
                if not pd or pd.get("tip") == "yok":
                    for cf in ciftler:
                        yazilanlar.pop((r, cf[0]), None)
                    continue
                if pd.get("tip") == "satir":
                    payda = satir_degeri(int(pd["satir"]))
                else:
                    pk = pd.get("kaynak") or s.get("kaynak") or kural.get("kaynak")
                    payda = topla(kitle(pk, pd.get("filtre"), ay), agirlik_fn(rapor["kaynaklar"][pk]))
                for cf in ciftler:
                    yaz(r, cf[0], v / payda if payda else 0)

    def pivot_doldur(p, kural):
        kaynak = rapor["kaynaklar"].get(kural.get("kaynak"))
        if not kaynak:
            return "Veri kaynağı seçilmemiş."
        satirlar = veri.get(kural["kaynak"]) or []
        agirlik = agirlik_fn(kaynak)
        dk = kaynak["donemKolonu"]
        sinirlar = grup_sinirlari(p)
        if len(sinirlar) != len(aylar):
            return f"Şablonda {len(sinirlar)} dönem sütunu var, pencere {len(aylar)} ay."

        pop = [s for s in satirlar if js_sayi(s.get(dk)) in aylar]
        pop = suz(pop, kural.get("filtre"))
        govde = [r for r in p["labelRows"] if r != p["gtRow"]] if p["gtRow"] else ozet_govde(sayfa, p)
        if kural.get("agac"):
            ag = agactan_bloklar(kural["agac"], rapor.get("gruplar"))
            if ag["hata"]:
                return ag["hata"]
            for n in ag["notlar"]:
                mesaj("bilgi", p["anahtar"], n)
            bloklar = ag["bloklar"]
        else:
            bloklar = kural.get("bloklar") or []
        ozel_var = bool(ozel_satirlari(kural))
        if not bloklar and not ozel_var:
            return "Tabloda satır tanımı yok."
        if not bloklar:
            basliklari_yaz(p)
            ozel_uygula(p, kural, govde, sinirlar)
            return True
        for bi0, b0 in enumerate(bloklar):
            if not b0.get("satir"):
                return f"{b0.get('ad') or 'Parça ' + str(bi0 + 1)}: satırlara gelecek alan seçilmedi, tablo yazılmadı."
            if b0.get("alt") and not b0["alt"].get("satir"):
                return f"{b0.get('ad') or 'Parça ' + str(bi0 + 1)}: alt satır alanı seçilmedi."
        def sablona_gore_dagit(bloklar, govde):
            """Sabit değerler şablonda kendi etiketlerinin satırına, değişken
            parçalar iki başlık arasındaki satırlara (ör. gruplanmış kod listeleri)."""
            def norm(v):
                return metin(v).upper()
            basliklar = {norm(x.get("etiket") or x["degerler"][0]) for b in bloklar if b.get("mod") == "sabit" for x in b.get("sabit") or []}
            i, dagitim, atlanan = 0, [], []
            for b in bloklar:
                if b.get("mod") == "sabit":
                    satirlar = []
                    for x in b.get("sabit") or []:
                        hedef, j = norm(x.get("etiket") or x["degerler"][0]), i
                        while j < len(govde) and norm(sayfa.deger(govde[j], p["labelCol"])) != hedef:
                            j += 1
                        if j >= len(govde):
                            atlanan.append(x.get("etiket") or x["degerler"][0])
                            satirlar.append(None)
                        else:
                            satirlar.append(govde[j])
                            i = j + 1
                    dagitim.append(satirlar)
                else:
                    bas = i
                    while i < len(govde) and norm(sayfa.deger(govde[i], p["labelCol"])) not in basliklar:
                        i += 1
                    dagitim.append(govde[bas:i])
            return {"dagitim": dagitim, "uyari": None, "kullanilmayan": govde[i:],
                    "bilgi": f"Şablonda satırı olmadığı için gösterilmeyen: {', '.join(atlanan)}." if atlanan else None}

        d = sablona_gore_dagit(bloklar, govde) if kural.get("yerlesim") == "sablon" else satirlari_dagit(bloklar, govde)
        if d.get("bilgi"):
            mesaj("bilgi", p["anahtar"], d["bilgi"])
        if d["uyari"]:
            mesaj("uyari", p["anahtar"], d["uyari"])
        if d["kullanilmayan"]:
            mesaj("bilgi", p["anahtar"], f"{len(d['kullanilmayan'])} şablon satırına hiçbir değer düşmedi, olduğu gibi bırakıldı.")

        # Sütunlar: her ay grubunda birden çok adet sütunu varsa (Onay / Red /
        # Toplam) her sütun kendi dilimini gösterir.
        sut = kural.get("sutunlar") if (kural.get("sutunlar") or {}).get("alan") else None
        dilimler = [{"etiket": x.get("etiket"), "kosul": {"kolon": sut["alan"], "op": "in", "degerler": x["degerler"]}}
                    for x in sut["degerler"]] if sut else []
        if sut and sut.get("toplam") is not False:
            dilimler.append({"etiket": "Toplam", "kosul": None, "toplam": True})
        if not sut:
            dilimler.append({"etiket": None, "kosul": None, "toplam": True})
        toplam_dilim = len(dilimler) - 1
        if sut:
            ornek0 = govde[0] if govde else p["gtRow"]
            sn = len([1 for _, t in sutunlari_siniflandir(sayfa, ornek0, sinirlar[0][0], sinirlar[0][1], kural.get("genelAdet")) if t == "adet"])
            if sn != len(dilimler):
                return (f"Şablonda her ay için {sn} adet sütunu var ama sütun alanında {len(dilimler)} sütun tanımlı"
                        + (" (Toplam dahil)" if sut.get("toplam") is not False else "") + ".")

        son_ay = aylar[-1]

        def sablon_gruplari(satir_nolari, alt):
            # Alt seviyenin sabit listesi varsa etiket oradan eşlenir, değerleri
            # oradan gelir (ör. "İnceleme" = birden çok alt kategori).
            if alt.get("sabit"):
                kume = {metin(x.get("etiket") or x["degerler"][0]).upper(): [metin(d) for d in x["degerler"]] for x in alt["sabit"]}
            else:
                kol = next((c for c in kaynak.get("kolonlar") or [] if c["ad"] == alt["satir"]), {}).get("degerler") or []
                kume = {metin(v).upper(): None for v in kol}
            gruplar = []
            for r in satir_nolari:
                e = sayfa.deger(r, p["labelCol"])
                if isinstance(e, str) and e.strip().upper() in kume and gruplar:
                    gruplar[-1]["cocuklar"].append({"satir": r, "etiket": e, "degerler": kume[e.strip().upper()]})
                else:
                    gruplar.append({"ust": r, "cocuklar": []})
            return gruplar

        def alt_havuz(bpop, b, gi, g):
            kolon = b["satir"]
            if gi.get("diger"):
                gosterilen = [v for x in g["girdiler"] if x.get("degerler") for v in x["degerler"]]
                return [s for s in bpop if metin(s.get(kolon)) not in g["haric"] and metin(s.get(kolon)) not in gosterilen]
            return [s for s in bpop if metin(s.get(kolon)) in gi["degerler"]]

        # Hangi değerin hangi satıra geleceği pencerenin tamamına (ya da son
        # aya) bakılarak BİR KEZ belirlenir.
        planlar = []
        for bi, b in enumerate(bloklar):
            bpop = suz(pop, b.get("filtre"))
            satir_nolari = d["dagitim"][bi]
            plan = []
            if b.get("alt") and b["alt"].get("mod") == "sablon":
                gruplar = sablon_gruplari(satir_nolari, b["alt"])
                g = girdileri_belirle(b, bpop, son_ay, dk, agirlik, len(gruplar), aylar)
                for i, gi in enumerate(g["girdiler"]):
                    plan.append({"satir": gruplar[i]["ust"], "girdi": gi, "cocuklar": [
                        {"satir": c["satir"], "girdi": {"bos": True} if gi.get("bos") else
                         {"etiket": c["etiket"], "degerler": c["degerler"] or [metin(c["etiket"])]}}
                        for c in gruplar[i]["cocuklar"]]})
            else:
                alt_yer = alt_satir_sayisi(b.get("alt")) if b.get("alt") else 0
                grup_boyu = 1 + alt_yer
                g = girdileri_belirle(b, bpop, son_ay, dk, agirlik, len(satir_nolari) // grup_boyu, aylar)
                for i, gi in enumerate(g["girdiler"]):
                    x = {"satir": satir_nolari[i * grup_boyu], "girdi": gi, "cocuklar": []}
                    if b.get("alt"):
                        ag2 = None if gi.get("bos") else girdileri_belirle(b["alt"], alt_havuz(bpop, b, gi, g), son_ay, dk, agirlik, alt_yer, aylar)
                        for j in range(alt_yer):
                            x["cocuklar"].append({"satir": satir_nolari[i * grup_boyu + 1 + j],
                                                  "girdi": ag2["girdiler"][j] if ag2 else {"bos": True}})
                    plan.append(x)
            # Şablona göre yerleşimde satırı bulunmayan sabit değer gösterilmez.
            plan = [x for x in plan if x["satir"] is not None]
            if b.get("mod") != "sabit" and g["adaySayisi"] > g["yaprakYer"]:
                mesaj("bilgi", p["anahtar"], f"{b.get('ad') or 'Parça ' + str(bi + 1)}: '{b['satir']}' kolonunda {g['adaySayisi']} farklı değer var, "
                      f"{g['yaprakYer']} satıra sığdı" + ("; kalanı 'Diğer'de." if (b.get("diger") or {}).get("aktif") else "; sığmayanlar gösterilmiyor."))
            if not bpop and pop:
                mesaj("uyari", p["anahtar"], f"{b.get('ad') or 'Parça ' + str(bi + 1)}: üst satırların değerleriyle süzülünce hiç kayıt kalmadı — değerlerin yazımını kontrol edin.")
            planlar.append({"blok": b, "plan": plan, "haric": g["haric"]})

        gt_satir = p["gtRow"]
        temizlenecek = []

        # Paketleme: kullanılan satırlar boşluksuz dizilir, başlıklar (sabit
        # değerler, Grand Total) biçimleriyle yukarı taşınır, taşındıkları satır
        # sıradan kod satırının biçimine döner, artanlar temizlenir.
        if kural.get("paketle") and p["gtRow"]:
            mantiksal = []
            for pl in planlar:
                for x in pl["plan"]:
                    if x["girdi"].get("bos"):
                        continue
                    mantiksal.append({"nesne": x, "baslik": pl["blok"].get("mod") == "sabit"})
                    mantiksal += [{"nesne": c, "baslik": False} for c in x["cocuklar"] if not c["girdi"].get("bos")]
            mantiksal.sort(key=lambda m: m["nesne"]["satir"])
            tum = [r for r in p["labelRows"] if r <= p["gtRow"]]
            if len(mantiksal) + 1 > len(tum):
                mesaj("uyari", p["anahtar"], f"Satırlar paketlenemedi: {len(mantiksal) + 1} satır gerekiyor, şablonda {len(tum)} yer var.")
            else:
                ilk_dinamik = next((d["dagitim"][bi][0] for bi, pl in enumerate(planlar)
                                    if pl["blok"].get("mod") != "sabit" and d["dagitim"][bi]), None)

                def stil_tasi(hedef, kaynak):
                    for c in range(p["labelCol"], p["rightBoundary"] + 1):
                        stiller[(hedef, c)] = (kaynak, c)
                gt_etiket = sayfa.deger(p["gtRow"], p["labelCol"])
                mantiksal.append({"gt": True, "satir": p["gtRow"], "baslik": True})
                for i, m in enumerate(mantiksal):
                    kaynak = m["satir"] if m.get("gt") else m["nesne"]["satir"]
                    hedef = tum[i]
                    if m["baslik"] and kaynak != hedef:
                        stil_tasi(hedef, kaynak)
                        if ilk_dinamik is not None:
                            stil_tasi(kaynak, ilk_dinamik)
                    if m.get("gt"):
                        gt_satir = hedef
                    else:
                        m["nesne"]["satir"] = hedef
                for pl in planlar:
                    pl["plan"] = [x for x in pl["plan"] if not x["girdi"].get("bos")]
                    for x in pl["plan"]:
                        x["cocuklar"] = [c for c in x["cocuklar"] if not c["girdi"].get("bos")]
                artan = tum[len(mantiksal):]
                temizlenecek += artan
                if gt_satir != p["gtRow"]:
                    mesaj("bilgi", p["anahtar"], f"Satırlar paketlendi: Grand Total R{p['gtRow']} yerine R{gt_satir}'e çekildi, {len(artan)} satır temizlendi.")
                    yaz(gt_satir, p["labelCol"], gt_etiket)

        # Boş satırlar: varsayılan etiket silinip 0 yazılır; "temizle" ise tüm
        # hücreler boşaltılır.
        if kural.get("bosSatir") == "temizle":
            for pl in planlar:
                for x in pl["plan"]:
                    if x["girdi"].get("bos"):
                        temizlenecek += [x["satir"]] + [c["satir"] for c in x["cocuklar"]]
                    else:
                        temizlenecek += [c["satir"] for c in x["cocuklar"] if c["girdi"].get("bos")]

        for pl in planlar:
            for x in pl["plan"]:
                yaz(x["satir"], p["labelCol"], None if x["girdi"].get("bos") else x["girdi"]["etiket"])
                for c in x["cocuklar"]:
                    yaz(c["satir"], p["labelCol"], None if c["girdi"].get("bos") else c["girdi"]["etiket"])
        basliklari_yaz(p)

        def deg(havuz, kolon, gi, kard):
            if gi.get("bos"):
                return 0
            if gi.get("diger"):
                t, es = topla(havuz, agirlik), 0
                for k in kard:
                    if k.get("degerler"):
                        es += topla([s for s in havuz if metin(s.get(kolon)) in k["degerler"]], agirlik)
                return max(0, t - es)
            return topla([s for s in havuz if metin(s.get(kolon)) in gi["degerler"]], agirlik)

        def ay_degerleri(ay_pop):
            liste, ust_toplam = [], 0
            for pl in planlar:
                b = pl["blok"]
                bm = [s for s in suz(ay_pop, b.get("filtre")) if metin(s.get(b["satir"])) not in pl["haric"]]
                blok_toplam = topla(bm, agirlik)
                # Düzleşen alt seviyede "üstündeki satır" = parça filtresindeki
                # kayıtlar (hariç tutulanlar dahil).
                blok_ham = topla(suz(ay_pop, b.get("filtre")), agirlik) if b.get("payda") == "blok_ham" else 0
                kardesler = [x["girdi"] for x in pl["plan"]]
                for x in pl["plan"]:
                    v = deg(bm, b["satir"], x["girdi"], kardesler)
                    if not b.get("derinlik"):
                        ust_toplam += v
                    liste.append({"satir": x["satir"], "adet": v, "payda": b.get("payda") or "grand_total",
                                  "blokToplam": blok_toplam, "blokHam": blok_ham})
                    if not x["cocuklar"]:
                        continue
                    gi = x["girdi"]
                    if gi.get("bos"):
                        alt_kaynak = []
                    elif gi.get("diger"):
                        alt_kaynak = [s for s in bm if not any(k.get("degerler") and metin(s.get(b["satir"])) in k["degerler"] for k in kardesler)]
                    else:
                        alt_kaynak = [s for s in bm if metin(s.get(b["satir"])) in gi["degerler"]]
                    alt_haric = [metin(h) for h in b["alt"].get("haric") or []]
                    alt_kaynak = [s for s in alt_kaynak if metin(s.get(b["alt"]["satir"])) not in alt_haric]
                    alt_kard = [c["girdi"] for c in x["cocuklar"]]
                    for c in x["cocuklar"]:
                        liste.append({"satir": c["satir"], "adet": deg(alt_kaynak, b["alt"]["satir"], c["girdi"], alt_kard),
                                      # Boş bir grubun alt satırlarında payda da 0'dır.
                                      "payda": "sifir" if gi.get("bos") else (b["alt"].get("payda") or "ust"),
                                      "ust": v, "blokToplam": topla(alt_kaynak, agirlik)})
            gt = ust_toplam if kural.get("grandTotal") == "gosterilen" else topla(ay_pop, agirlik)
            return {"satirlar": liste, "gt": gt}

        hesap = []
        for ay in aylar:
            ay_pop = [s for s in pop if js_sayi(s.get(dk)) == ay]
            hesap.append([ay_degerleri(suz(ay_pop, [dl["kosul"]]) if dl["kosul"] else ay_pop) for dl in dilimler])
        gt_pencere = [sum(h[k]["gt"] for h in hesap) for k in range(len(dilimler))]
        satir_pencere = [sum(h[toplam_dilim]["satirlar"][j]["adet"] for h in hesap) for j in range(len(hesap[0][toplam_dilim]["satirlar"]))]
        pencere_mi = kural.get("oranPaydasi") == "pencere"

        for ai, ay in enumerate(aylar):
            sinir = sinirlar[ai]
            siniflar = sutunlari_siniflandir(sayfa, govde[0] if govde else p["gtRow"], sinir[0], sinir[1], kural.get("genelAdet"))
            ciftler = adet_oran_eslestir(siniflar)
            adet_sutunlari = [c for c, t in siniflar if t == "adet"]

            def dilim_of(adet_col):
                return adet_sutunlari.index(adet_col) if sut else 0
            h = hesap[ai]
            for j in range(len(h[0]["satirlar"])):
                for c in adet_sutunlari:
                    sd = h[dilim_of(c)]["satirlar"][j]
                    yaz(sd["satir"], c, sd["adet"])
                for cf in ciftler:
                    k = dilim_of(cf[1])
                    sd = h[k]["satirlar"][j]
                    if sut and not dilimler[k].get("toplam"):
                        op = sut.get("oranPayda") or "satir"
                        if op == "satir_pencere":
                            payda = satir_pencere[j]
                        elif op == "grand_total":
                            payda = h[k]["gt"]
                        else:
                            payda = h[toplam_dilim]["satirlar"][j]["adet"]
                    elif sd["payda"] == "yok":
                        continue
                    elif sd["payda"] == "sifir":
                        payda = 0
                    elif sd["payda"] == "blok":
                        payda = sd["blokToplam"]
                    elif sd["payda"] == "ust":
                        payda = sd["ust"]
                    elif sd["payda"] == "blok_ham":
                        payda = sd["blokHam"]
                    else:
                        payda = gt_pencere[k] if pencere_mi else h[k]["gt"]
                    yaz(sd["satir"], cf[0], sd["adet"] / payda if payda else 0)
            if not gt_satir:
                continue
            for c in adet_sutunlari:
                yaz(gt_satir, c, h[dilim_of(c)]["gt"])
            for cf in ciftler:
                k = dilim_of(cf[1])
                gt = h[k]["gt"]
                if sut and not dilimler[k].get("toplam"):
                    op = sut.get("oranPayda") or "satir"
                    payda = gt_pencere[toplam_dilim] if op == "satir_pencere" else gt if op == "grand_total" else h[toplam_dilim]["gt"]
                else:
                    payda = gt_pencere[k] if pencere_mi else gt
                yaz(gt_satir, cf[0], gt / payda if payda else 0)

        for r in temizlenecek:
            for c in range(p["labelCol"], p["rightBoundary"] + 1):
                yaz(r, c, None)
        ozel_uygula(p, kural, govde, sinirlar)
        return True

    def toplam_sutunu_yaz(p, kural):
        """TOTAL sütunu: satırın dönem adetlerinin toplamı; oranı Grand Total
        toplamına göre."""
        tg = toplam_grubu(sayfa, p)
        if not tg:
            mesaj("uyari", p["anahtar"], "Toplam sütunu istendi ama şablonda dönem sütunlarının sağında başlıklı bir sütun grubu yok.")
            return
        govde = [r for r in p["labelRows"] if r != p["gtRow"]] if p["gtRow"] else ozet_govde(sayfa, p)
        ornek = govde[0] if govde else p["gtRow"]
        adetler = [[c for c, t in sutunlari_siniflandir(sayfa, ornek, g[0], g[1], kural.get("genelAdet")) if t == "adet"]
                   for g in grup_sinirlari(p)[:len(aylar)]]
        siniflar = sutunlari_siniflandir(sayfa, ornek, tg[0], tg[1], kural.get("genelAdet"))
        hedef = [c for c, t in siniflar if t == "adet"]
        ortalama = {govde[i] for i, o in ozel_satirlari(kural).items()
                    if i < len(govde) and (o.get("olcu") or {}).get("tip") == "ortalama"}
        toplamlar = {}
        for r in p["labelRows"]:
            if r in ortalama:
                continue
            t = {}
            for j, c in enumerate(hedef):
                top, tamam = 0, True
                for a in adetler:
                    v = yazilanlar.get((r, a[j])) if j < len(a) else None
                    if not isinstance(v, (int, float)) or isinstance(v, bool):
                        tamam = False
                        break
                    top += v
                if not tamam:
                    continue
                t[c] = top
                yaz(r, c, top)
            toplamlar[r] = t
        gt = toplamlar.get(p["gtRow"]) if p["gtRow"] else None
        if not gt:
            return
        for oc, ac in adet_oran_eslestir(siniflar):
            payda = gt.get(ac)
            if payda is None:
                continue
            for r in p["labelRows"]:
                pay = toplamlar.get(r, {}).get(ac)
                if pay is not None:
                    yaz(r, oc, pay / payda if payda else 0)

    for p in paneller:
        kural = kurallar.get(p["anahtar"])
        if not kural:
            basliklari_yaz(p)
            sonuc["kuralsiz"].append(p["anahtar"])
            continue
        try:
            r = pivot_doldur(p, kural)
            if r is True and kural.get("toplamSutunu"):
                toplam_sutunu_yaz(p, kural)
            if r is True:
                sonuc["dolu"].append(p["anahtar"])
            else:
                sonuc["yazilamadi"].append(p["anahtar"])
                mesaj("uyari", p["anahtar"], r)
        except Exception as e:  # noqa: BLE001 — bir tablo hatası raporu durdurmaz
            sonuc["hata"].append(p["anahtar"])
            mesaj("hata", p["anahtar"], f"Tablo doldurulurken hata: {e}")

    anahtarlar = {p["anahtar"] for p in paneller}
    for k in kurallar:
        if k not in anahtarlar:
            mesaj("uyari", k, "Bu kuralın şablonda karşılığı yok — başlık değişmiş olabilir.")
    return {"yazilanlar": yazilanlar, "sonuc": sonuc, "mesajlar": mesajlar, "stiller": stiller}


# %% Rapor üretimi | `rapor_uret` her şeyi birleştirir: gereken kaynakları `_read_source` ile okur, en güncel dönemi bulur, şablonu açar, sayfaları doldurur, raporda olmayan sayfaları siler ve dosyayı kaydeder.


def gerekli_kaynaklar(rapor, sayfa_adlari):
    """Yalnızca bu notebook'un sayfalarında kullanılan kaynaklar okunur."""
    ids = []

    def ekle(k):
        if k and k in rapor["kaynaklar"] and k not in ids:
            ids.append(k)
    for ad in sayfa_adlari:
        for kural in ((rapor["sayfalar"].get(ad) or {}).get("paneller") or {}).values():
            ekle(kural.get("kaynak"))
            ozel = kural.get("ozelSatirlar") or {i: s for i, s in enumerate(kural.get("satirlar") or [])}
            for s in ozel.values():
                ekle(s.get("kaynak"))
                ekle((s.get("payda") or {}).get("kaynak"))
    return ids


def gerekli_kolonlar(rapor, kaynak_id, sayfa_adlari):
    """Veriden yalnızca kuralların baktığı kolonlar alınır (bellek ve hız)."""
    k = rapor["kaynaklar"][kaynak_id]
    kolonlar = {k["donemKolonu"]}
    if k.get("adetKolonu"):
        kolonlar.add(k["adetKolonu"])
    gruplar = {g["ad"]: g for g in rapor.get("gruplar") or [] if g["kaynak"] == kaynak_id}

    def kosullar(liste):
        for f in liste or []:
            kolonlar.add(f["kolon"])
            if f.get("istisna"):
                kolonlar.add(f["istisna"]["kolon"])
    for ad in sayfa_adlari:
        for kural in ((rapor["sayfalar"].get(ad) or {}).get("paneller") or {}).values():
            ozel = kural.get("ozelSatirlar") or {i: s for i, s in enumerate(kural.get("satirlar") or [])}
            if kural.get("kaynak") == kaynak_id:
                kosullar(kural.get("filtre"))
                for L in (kural.get("agac") or {}).get("seviyeler") or []:
                    if L.get("satir"):
                        kolonlar.add(L["satir"])
                if (kural.get("sutunlar") or {}).get("alan"):
                    kolonlar.add(kural["sutunlar"]["alan"])
            for s in ozel.values():
                if (s.get("kaynak") or kural.get("kaynak")) == kaynak_id:
                    kosullar(s.get("filtre"))
                    if (s.get("olcu") or {}).get("kolon"):
                        kolonlar.add(s["olcu"]["kolon"])
                pd = s.get("payda") or {}
                if pd.get("tip") == "filtre" and (pd.get("kaynak") or s.get("kaynak") or kural.get("kaynak")) == kaynak_id:
                    kosullar(pd.get("filtre"))
    # Grup (türetilmiş) kolonların yerine ham kolonları okunur. Koşullu bir
    # alan birden çok kolona bakar (ör. Accept = karar kategorisi YA DA nihai
    # karar); yalnızca `kolon` alınırsa diğer koşul sessizce hep yanlış çıkar.
    # Bir grup başka bir gruba dayanabileceği için açılım tekrarlanır.
    def ham(ad, gorulen):
        g = gruplar.get(ad)
        if g is None or ad in gorulen:
            return {ad}
        gorulen = gorulen | {ad}
        adlar = {g["kolon"]} if g.get("kolon") else set()
        for kr in g.get("kurallar") or []:
            for secenek in kr.get("secenekler") or []:
                for f in secenek:
                    adlar.add(f["kolon"])
                    if f.get("istisna"):
                        adlar.add(f["istisna"]["kolon"])
        return set().union(*(ham(a, gorulen) for a in adlar)) if adlar else set()
    return sorted(set().union(*(ham(a, frozenset()) for a in kolonlar)))


def _donem_serisi(df, kolon):
    import pandas as pd
    return pd.to_numeric(df[kolon], errors="coerce")


def veriyi_hazirla(df, rapor, kaynak_id, sayfa_adlari, aylar):
    """DataFrame → motorun kullandığı kayıt listesi (yalnızca pencere ve gereken kolonlar)."""
    import pandas as pd
    k = rapor["kaynaklar"][kaynak_id]
    dk = k["donemKolonu"]
    if dk not in df.columns:
        raise RuntimeError(f"'{k['ad']}' verisinde dönem kolonu '{dk}' yok. Gelen kolonlar: {list(df.columns)}")
    istenen = gerekli_kolonlar(rapor, kaynak_id, sayfa_adlari)
    eksik = [c for c in istenen if c not in df.columns]
    if eksik:
        print(f"[UYARI] '{k['ad']}' verisinde şu kolon(lar) yok, boş sayılacak: {eksik}. "
              f"Gelen kolonlar: {list(df.columns)}")
    var_olan = [c for c in istenen if c in df.columns]
    secili = df.loc[_donem_serisi(df, dk).isin(aylar), var_olan]
    secili = secili.astype(object).where(pd.notna(secili), None)
    kayitlar = secili.to_dict("records")
    gruplari_uygula(kayitlar, [g for g in rapor.get("gruplar") or [] if g["kaynak"] == kaynak_id])
    print(f"[BİLGİ] {k['ad']}: {len(df)} satırdan dönem penceresindeki {len(kayitlar)} satır kullanılıyor "
          f"({len(var_olan)} kolon).")
    return kayitlar


def rapor_uret(rapor, sayfa_adlari, kaynak_oku, sablon_yukle, hedef_donem=None, cikti="rapor.xlsx",
               pencere_ister=None):
    """pencere_ister(kaynak_id): sorgusu raporun aylarını isteyen kaynaklar
    (ör. ay sonu tarihli bir sorgu) dönem bulunduktan sonra, o aylarla okunur."""
    from pathlib import Path

    kaynak_idleri = gerekli_kaynaklar(rapor, sayfa_adlari)
    sonra = [k for k in kaynak_idleri if pencere_ister and pencere_ister(k)]
    ham = {}

    def oku(kid, aylar=None):
        ham[kid] = kaynak_oku(kid, aylar=aylar) if aylar else kaynak_oku(kid)
        print(f"[BİLGİ] {rapor['kaynaklar'][kid]['ad']}: {len(ham[kid])} satır x {len(ham[kid].columns)} kolon okundu.")
    for kid in kaynak_idleri:
        if kid not in sonra:
            oku(kid)

    if hedef_donem is None:
        adaylar = []
        for kid, df in ham.items():
            dk = rapor["kaynaklar"][kid]["donemKolonu"]
            if dk in df.columns:
                adaylar += [int(v) for v in _donem_serisi(df, dk).dropna().unique() if float(v).is_integer() and donem_mi(int(v))]
        if not adaylar:
            raise RuntimeError("Verilerde geçerli bir dönem (YYYYMM) bulunamadı; HEDEF_DONEM'i elle verin.")
        hedef_donem = max(adaylar)
    print(f"[BİLGİ] Hedef dönem: {hedef_donem}")

    okuma = sablon_yukle()
    yazma = sablon_yukle()
    eksik_sayfa = [a for a in sayfa_adlari if a not in okuma.sheetnames]
    if eksik_sayfa:
        raise RuntimeError(f"Şablonda şu sayfa(lar) yok: {eksik_sayfa}. Şablondaki sayfalar: {okuma.sheetnames}")

    modeller, pencereler = {}, {}
    for ad in sayfa_adlari:
        model = SayfaModeli(okuma[ad])
        paneller = panelleri_kesfet(model)
        n = len(sablon_aylari(model, paneller)) or 3
        modeller[ad] = (model, paneller)
        pencereler[ad] = pencere(hedef_donem, n)
    tum_aylar = sorted({a for p in pencereler.values() for a in p})
    for kid in sonra:
        oku(kid, tum_aylar)
    ham = {k: ham[k] for k in kaynak_idleri}
    veri = {kid: veriyi_hazirla(df, rapor, kid, sayfa_adlari, tum_aylar) for kid, df in ham.items()}

    for ad in sayfa_adlari:
        model, paneller = modeller[ad]
        aylar = pencereler[ad]
        sonuc = sayfa_doldur(model, paneller, rapor["sayfalar"].get(ad), rapor, veri, aylar)
        ws = yazma[ad]
        yazilamayan = 0
        # Paketlenen tablolarda taşınan satırların biçimi şablondan kopyalanır.
        from copy import copy as _kopya
        for (r, c), (r0, c0) in sonuc["stiller"].items():
            hedef = ws.cell(row=r, column=c)
            if type(hedef).__name__ != "MergedCell":
                hedef._style = _kopya(okuma[ad].cell(row=r0, column=c0)._style)
        for (r, c), v in sonuc["yazilanlar"].items():
            hucre = ws.cell(row=r, column=c)
            if type(hucre).__name__ == "MergedCell":
                yazilamayan += 1
                continue
            hucre.value = v
        s = sonuc["sonuc"]
        print(f"\n[BİLGİ] {ad} ({aylar[0]}–{aylar[-1]}): {len(s['dolu'])} tablo dolu"
              + (f", {len(s['kuralsiz'])} tablonun kuralı yok (atlandı): {s['kuralsiz']}" if s["kuralsiz"] else "")
              + (f", {len(s['yazilamadi']) + len(s['hata'])} tablo yazılamadı" if s["yazilamadi"] or s["hata"] else ""))
        for seviye, panel, m in sonuc["mesajlar"]:
            etiket = {"bilgi": "[BİLGİ]", "uyari": "[UYARI]", "hata": "[HATA]"}[seviye]
            print(f"{etiket} {ad} / {panel}: {m}")
        if yazilamayan:
            print(f"[UYARI] {ad}: {yazilamayan} değer birleşik hücrenin içine düştüğü için yazılamadı.")

    # Raporda olmayan sayfalar çıktıdan silinir.
    for ad in list(yazma.sheetnames):
        if ad not in sayfa_adlari:
            del yazma[ad]
    cikti = Path(cikti)
    yazma.save(cikti)
    print(f"\nRapor oluşturuldu: {cikti.resolve()}")
    return cikti
