# -*- coding: utf-8 -*-
"""Kaizen sisteminin tasarim sistemi -- tek kaynak.

Buradaki hex degerleri `kaynak/vba/modTasarim.bas` icindeki sabitlerle
birebir ayni olmak zorundadir. kur.py uretimden once ikisini karsilastirir;
uyusmazlik varsa uretim durur.
"""

# --------------------------------------------------------------------------
# Kurumsal palet (Is Bankasi lacivert agirlikli)
# --------------------------------------------------------------------------
RENK = {
    "ANA_LACIVERT":   "002D62",   # masthead, birincil dugme, KPI rakami
    "DERIN_LACIVERT": "001F44",   # koyu vurgu, kapak zemini
    "KURUMSAL_MAVI":  "1B5FAA",   # ikincil vurgu, grafik 2. seri
    "ACIK_MAVI":      "D6E4F0",   # secili satir, bilgi kutusu, seritler
    "SOLUK_MAVI":     "9FB6D4",   # masthead ikincil yazi
    "BUZ_ZEMIN":      "F4F7FB",   # sayfa zemini
    "BEYAZ":          "FFFFFF",   # kart zemini, giris alani
    "CIZGI_GRI":      "D9DEE6",   # kart kenarligi, ince tablo cizgisi
    "GOLGE_GRI":      "C6CDD6",   # kart sag/alt kenari (derinlik hissi)
    "ALAN_CIZGI":     "B7C4D6",   # form alani alt kenarligi
    "METIN_KOYU":     "1A2433",   # govde metni
    "METIN_GRI":      "5A6572",   # ikincil etiket
    "ONAY_ZEMIN":     "D2EAD9",   # basarili gonderim bandi
    "ONAY_YAZI":      "1E7B4D",
    "UYARI_ZEMIN":    "F6DADA",
    "UYARI_YAZI":     "9A2E2E",
}

# --------------------------------------------------------------------------
# Durum renkleri -- (zemin, yazi)
# --------------------------------------------------------------------------
DURUM_RENK = {
    "Yeni":                 ("D6E4F0", "002D62"),
    "Değerlendirmede":      ("FCEBCF", "8A5A00"),
    "Planlandı":            ("E4DBF5", "4B3A8C"),
    "Pilot Uygulamada":     ("D5EEF0", "0E6470"),
    "Ölçümleniyor":         ("DDEBD9", "3A6B35"),
    "Standartlaştırıldı":   ("D2EAD9", "1E7B4D"),
    "Beklemede":            ("EDEFF2", "5A6572"),
    "Reddedildi":           ("F6DADA", "9A2E2E"),
}

# --------------------------------------------------------------------------
# Oncelik sinifi renkleri -- (zemin, yazi)
# --------------------------------------------------------------------------
ONCELIK_RENK = {
    "Hızlı Kazanım":        ("D2EAD9", "1E7B4D"),
    "Büyük Proje":          ("D6E4F0", "1B5FAA"),
    "Doldurma İşi":         ("FCEBCF", "C77700"),
    "Değerlendirme Dışı":   ("EDEFF2", "8A93A1"),
}

# --------------------------------------------------------------------------
# Tipografi
# --------------------------------------------------------------------------
FONT_AILE = "Segoe UI"
FONT_BASLIK_AILE = "Segoe UI Semibold"

# --------------------------------------------------------------------------
# Urun adi -- her sayfanin ustundeki lacivert bantta solda yazar.
# --------------------------------------------------------------------------
URUN_ADI = "XJ Birimi  |  Proje Öneri Formu"

PT = {
    "SAYFA_BASLIK": 18,
    "MASTHEAD_ALT": 10,
    "BOLUM_BASLIK": 12,
    "GOVDE":        10.5,
    "ETIKET":        9,
    "KPI_RAKAM":    26,
    "KPI_ETIKET":    9,
    "KPI_ALT":       9,
    "TABLO_BASLIK": 10,
}

# --------------------------------------------------------------------------
# Olculer
# --------------------------------------------------------------------------
OLCU = {
    "MASTHEAD_YUKSEKLIK": 22.0,   # masthead satir yuksekligi (2 satir)
    "SATIR_NORMAL":       16.5,
    "SATIR_BOSLUK":        6.0,
    "KART_YUKSEKLIK":     18.0,
    "SUTUN_DAR":           2.2,   # kenar bosluk sutunu
    "SUTUN_ETIKET":       22.0,
    "SUTUN_ALAN":         46.0,
}

# Dugme olculeri (nokta cinsinden; COM tarafinda kullanilir)
DUGME = {
    "GENISLIK":  148.0,
    "YUKSEKLIK":  34.0,
    "KOSE_ORAN":   0.18,
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
    for sinif, (zemin, yazi) in ONCELIK_RENK.items():
        duz[f"ONCELIK_{sinif}_ZEMIN"] = zemin
        duz[f"ONCELIK_{sinif}_YAZI"] = yazi
    return duz
