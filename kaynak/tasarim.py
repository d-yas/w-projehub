# -*- coding: utf-8 -*-
"""Proje Öneri sisteminin tasarim sistemi -- tek kaynak.

Buradaki hex degerleri `kaynak/vba/modTasarim.bas` icindeki sabitlerle
birebir ayni olmak zorundadir. kur.py uretimden once ikisini karsilastirir;
uyusmazlik varsa uretim durur.

Palet iki katmandan olusur:
  * MARKA  -- lacivert ve mavi. Kimlik burada; degistirilmesi kurumsal karardir.
  * NOTR   -- zemin, cizgi ve metin gri basamaklari. Arayuzun "havasi" burada;
              acik zemin + hairline cizgi + genis bosluk, ekranin hucre
              izgarasi gibi degil bir uygulama gibi gorunmesini saglar.
"""

# --------------------------------------------------------------------------
# Kurumsal palet
# --------------------------------------------------------------------------
RENK = {
    # --- Marka -----------------------------------------------------------
    "ANA_LACIVERT":   "002D62",   # masthead, birincil dugme
    "DERIN_LACIVERT": "001F44",   # golge rengi, sekme rengi
    "KURUMSAL_MAVI":  "1F6FEB",   # vurgu cizgisi, sessiz dugme, trend grafigi
    "ACIK_MAVI":      "E8F1FD",   # bilgi kutusu zemini
    "SOLUK_MAVI":     "A8C0E0",   # masthead ikincil yazi

    # --- Notr basamaklar -------------------------------------------------
    "BUZ_ZEMIN":      "F5F7FA",   # sayfa zemini
    "BEYAZ":          "FFFFFF",   # kart ve giris alani zemini
    "CIZGI_INCE":     "EEF2F7",   # tablo satir ayraci, ince ayrac
    "CIZGI_GRI":      "E6EBF2",   # kart kenari, alan kenari
    "GOLGE_GRI":      "DCE3EC",   # kart alt kenari (hafif derinlik)
    "ALAN_CIZGI":     "C9D4E3",   # form alani alt kenarligi
    "METIN_KOYU":     "0F1B2E",   # govde metni
    "METIN_GRI":      "64748B",   # ikincil metin, etiket
    "METIN_SOLUK":    "94A3B8",   # ucuncul metin: ipucu, kart alt notu

    # --- Durum bantlari ve vurgular --------------------------------------
    "ONAY_ZEMIN":     "E3F5EA",
    "ONAY_YAZI":      "15803D",
    "UYARI_ZEMIN":    "FDE8E8",
    "UYARI_YAZI":     "B42318",
    "VURGU_KEHRIBAR": "B45309",   # "degerlendirme bekleyen" gostergesi
}

# --------------------------------------------------------------------------
# Durum renkleri -- (zemin, yazi)
#
# Zemin tonlari rozetlerde, YAZI tonlari hem rozet yazisinda hem de durum
# dagilimi grafiginin cubuklarinda kullanilir. Bu yuzden yazi tonlari
# birbirinden ayirt edilebilir olmak zorundadir.
# --------------------------------------------------------------------------
DURUM_RENK = {
    "Yeni":                 ("E8F1FD", "1D4ED8"),
    "Değerlendirmede":      ("FEF3C7", "92400E"),
    "Planlandı":            ("EDE9FE", "5B21B6"),
    "Pilot Uygulamada":     ("CFFAFE", "0E7490"),
    "Ölçümleniyor":         ("E7F6EC", "3F7A47"),
    "Standartlaştırıldı":   ("DCFCE7", "15803D"),
    "Beklemede":            ("F1F5F9", "64748B"),
    "Reddedildi":           ("FDE8E8", "B42318"),
}

# --------------------------------------------------------------------------
# Tipografi
# --------------------------------------------------------------------------
FONT_AILE = "Segoe UI"
FONT_BASLIK_AILE = "Segoe UI Semibold"

# ==========================================================================
#  BIRIM ADI -- degistirilecek tek yer.
#
#  Her sayfanin ustundeki lacivert bantta solda yazar. Kendi biriminizin
#  adini yazip "python kur.py" calistirin; butun ekranlarda degisir.
#  (Ekran sifreleri icin: kaynak/vba/modAyar.bas -- TALIMATNAME.md madde 1.)
# ==========================================================================
BIRIM_ADI = "XJ Birimi"

URUN_ADI = BIRIM_ADI + "  |  Proje Öneri Formu"

# Punto olcegi. Ekranlarda gecen HER punto degeri buradan gelir; sabit
# kodlanmis punto birakmayin, aksi halde olcek zamanla dagilir.
PT = {
    "MASTHEAD_BASLIK": 15,
    "MASTHEAD_ALT":     9.5,
    "EKRAN_BASLIK":    22,     # giris ekranlarindaki buyuk baslik
    "EKRAN_ALT":       11,
    "BOLUM_BASLIK":    11.5,
    "BOLUM_ALT":        9,
    "GOVDE":           10.5,
    "ETIKET":           8.5,   # mikro etiket (buyuk harf)
    "IPUCU":            8.5,
    "NOT":              9,
    "OZET":             9.5,
    "KPI_RAKAM":       26,
    "KPI_ETIKET":       8.5,
    "KPI_ALT":          8.5,
    "TABLO_BASLIK":     9.5,
    "GRAFIK":           8.5,   # eksen, etiket ve veri etiketi
    "DUGME":           10.5,
}

# --------------------------------------------------------------------------
# Olculer (satir yuksekligi punto cinsindendir)
# --------------------------------------------------------------------------
OLCU = {
    "MASTHEAD_UST":    14.0,   # bandin ust satiri
    "MASTHEAD_ALT":    28.0,   # bandin alt satiri (metin bu ikisinde ortalanir)
    "MASTHEAD_VURGU":   2.5,   # bandin altindaki ince mavi cizgi
    "SATIR_NORMAL":    18.0,
    "SATIR_BOSLUK":     6.0,
    "SATIR_TABLO":     20.0,
    "SATIR_ALAN":      24.0,   # tek satirlik form alani
}

# Dugme olculeri (nokta cinsinden; COM tarafinda kullanilir)
DUGME = {
    "GENISLIK":       150.0,
    "YUKSEKLIK":       36.0,
    "KOSE_ORAN":        0.22,
    "ARALIK":          10.0,   # yan yana dugmeler arasindaki bosluk
    "GOLGE_BULANIK":    8.0,
    "GOLGE_KAYDIRMA":   2.0,
    "GOLGE_SEFFAF":     0.82,
}

# Pano kartlari ve panelleri -- COM asamasinda eklenen sekiller.
#
# Excel hucrelere yuvarlak kose ve yumusak golge veremez, sekiller verebilir;
# ama sekiller HER ZAMAN hucrelerin ustunde cizilir. Bu yuzden kartin metni de
# sekle tasinir (bkz. uret_yonetim.com_ek_islem).
KART = {
    "KOSE_ORAN":        0.13,
    "IC_BOSLUK":       14.0,   # sekil ici yatay kenar boslugu
    "YATAY_INSET":      0.0,   # sekil hucre blogunu TAM kaplar (bkz. asagi)
    "GOLGE_BULANIK":   12.0,
    "GOLGE_KAYDIRMA":   3.0,
    "GOLGE_SEFFAF":     0.88,
    "CIZGI_KALINLIK":   0.75,
}


def rgb_long(hex_kod: str) -> int:
    """Hex rengi VBA'nin RGB() Long degerine cevirir (BGR sirasi)."""
    h = hex_kod.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + (g << 8) + (b << 16)


def tum_hexler() -> dict:
    """Tutarlilik kontrolu icin duz bir ad -> hex sozlugu uretir."""
    duz = dict(RENK)
    for durum, (zemin, yazi) in DURUM_RENK.items():
        duz[f"DURUM_{durum}_ZEMIN"] = zemin
        duz[f"DURUM_{durum}_YAZI"] = yazi
    return duz
