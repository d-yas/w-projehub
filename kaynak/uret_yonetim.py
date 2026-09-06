# -*- coding: utf-8 -*-
"""KaizenYonetim.xlsm -- Kaizen ekibinin konsol kitabinin sayfa kurulumu.

Bes ekran: Giris, Konsol, Degerlendirme, Pano, Rapor.
Ayrica iki gizli calisma sayfasi: Veri (konsolide tablo) ve PanoVeri
(grafiklerin kaynak araliklari).
"""

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import uret_ortak as u
from tasarim import DURUM_RENK, FONT_BASLIK_AILE, ONCELIK_RENK, PT, RENK

SAYFA_GIRIS = "Giriş"
SAYFA_KONSOL = "Konsol"
SAYFA_DEGERLENDIRME = "Değerlendirme"
SAYFA_PANO = "Pano"
SAYFA_RAPOR = "Rapor"
SAYFA_VERI = "Veri"
SAYFA_PANOVERI = "PanoVeri"

# --- Konsol duzeni (modKonsolide.bas'taki sabitlerle ayni olmali) --------
# Sutun genislikleri en uzun degerlerine gore olculmustur: oneri numarasi
# ("ON-260904-A7K") ve oncelik sinifi ("Değerlendirme Dışı") kirpilmamalidir.
KONSOL_SUTUNLAR = [2.2, 17.5, 12.0, 22.0, 46.0, 19.0,
                   6.5, 6.5, 22.0, 11.0, 13.0, 2.2]
KONSOL_BASLIK_SATIR = 8
KONSOL_ILK_SATIR = 9
KONSOL_STIL_SATIR = 500          # bu kadar satir onceden bicimlendirilir
KONSOL_ILK_SUTUN = 2             # B
KONSOL_SON_SUTUN = 11            # K

# --- Degerlendirme duzeni ------------------------------------------------
DEG_SUTUNLAR = [2.2, 24.0, 24.0, 24.0, 20.0, 26.0, 2.2]
DEG_GECMIS_BASLIK = 33
DEG_GECMIS_ILK = 34              # modDegerlendirme.GECMIS_ILK_SATIR
DEG_GECMIS_ADET = 14

# --- Pano duzeni ---------------------------------------------------------
PANO_SUTUNLAR = [2.2, 20.0, 20.0, 2.0, 20.0, 20.0, 2.0, 20.0, 20.0, 2.2]
PANO_KART_SUTUNLARI = [(2, 3), (5, 6), (8, 9)]     # (B,C) (E,F) (H,I)
PANO_GRAFIK_SUTUNLARI = [(2, 3), (5, 9)]           # (B,C) ve (E..I)
PANO_GRAFIK_UST = 23
PANO_GRAFIK_ALT = 38

RAPOR_SUTUNLAR = [2.2, 26.0, 26.0, 26.0, 26.0, 2.2]


def _ad(ad, sayfa, sutun, satir):
    return DefinedName(ad, attr_text=f"'{sayfa}'!${get_column_letter(sutun)}${satir}")


# ==========================================================================
#  Giris
# ==========================================================================
def _giris(wb):
    ws = wb.create_sheet(SAYFA_GIRIS)
    u.sayfa_hazirla(ws, [2.2, 24.0, 20.0, 20.0, 20.0, 14.0, 2.2], son_satir=30)
    u.masthead(ws, 7, "Giriş")

    u.bosluk(ws, 4, 18.0)
    b = u.birlestir(ws, 5, 2, 5, 6)
    b.value = "Kaizen Yönetim Konsolu"
    b.font = u.yazi(RENK["ANA_LACIVERT"], 20, kalin=True, aile=FONT_BASLIK_AILE)
    b.alignment = u.hiza("left", "center")
    ws.row_dimensions[5].height = 30.0

    a = u.birlestir(ws, 6, 2, 6, 6)
    a.value = "Öneriler burada değerlendirilir, önceliklendirilir ve raporlanır."
    a.font = u.yazi(RENK["METIN_GRI"], 11)
    ws.row_dimensions[6].height = 20.0

    u.bosluk(ws, 7, 14.0)
    u.bilgi_kutusu(
        ws, 8, 2, 9, 6,
        "Bu ekran yalnızca Kaizen ekibi içindir. Ekranda gördüğünüz her şey "
        "ortak klasördeki dosyalardan yeniden üretilir; bu dosya bozulsa bile "
        "veri kaybolmaz.",
    )
    ws.row_dimensions[8].height = 20.0
    ws.row_dimensions[9].height = 20.0

    u.bosluk(ws, 10, 44.0)      # "Konsola Gir" dugmesi
    u.bosluk(ws, 11, 12.0)

    n = u.birlestir(ws, 12, 2, 12, 6)
    n.value = ("Dosya açıldığında üstte bir güvenlik uyarısı çıkarsa "
               "“İçeriği Etkinleştir” düğmesine basın.")
    n.font = u.yazi(RENK["METIN_GRI"], 9, italik=True)

    ws.sheet_state = "visible"
    return ws


# ==========================================================================
#  Konsol
# ==========================================================================
def _konsol(wb):
    ws = wb.create_sheet(SAYFA_KONSOL)
    son_sutun = u.sayfa_hazirla(ws, KONSOL_SUTUNLAR,
                                son_satir=KONSOL_ILK_SATIR + KONSOL_STIL_SATIR,
                                baslik_gizle=False)
    u.masthead(ws, son_sutun, "Konsol")

    u.bosluk(ws, 4, 8.0)
    u.bosluk(ws, 5, 40.0)       # dugme seridi
    u.bosluk(ws, 6, 6.0)

    ozet = u.birlestir(ws, 7, KONSOL_ILK_SUTUN, 7, KONSOL_SON_SUTUN)
    ozet.font = u.yazi(RENK["METIN_GRI"], 9.5)
    ozet.alignment = u.hiza("left", "center")
    ozet.value = "Önerileri görmek için “Önerileri Yenile” düğmesine basın."
    ws.row_dimensions[7].height = 18.0
    wb.defined_names.add(_ad("knsl_ozet", SAYFA_KONSOL, KONSOL_ILK_SUTUN, 7))

    u.tablo_basligi(ws, KONSOL_BASLIK_SATIR, KONSOL_ILK_SUTUN, KONSOL_SON_SUTUN,
                    ["Öneri No", "Tarih", "Gönderen", "Öneri Başlığı",
                     "Durum", "Etki", "Efor", "Öncelik",
                     "Yıllık Saat", "Yıllık TL"])

    son = KONSOL_ILK_SATIR + KONSOL_STIL_SATIR - 1
    u.tablo_govde_stili(ws, KONSOL_ILK_SATIR, son, KONSOL_ILK_SUTUN, KONSOL_SON_SUTUN)

    # Sayisal sutunlar: etki/efor ortali, kazanimlar saga hizali.
    for sutun, bicim in ((7, "0"), (8, "0"), (10, "#,##0.0"), (11, "#,##0")):
        for r in range(KONSOL_ILK_SATIR, son + 1):
            h = ws.cell(row=r, column=sutun)
            h.number_format = bicim
            h.alignment = u.hiza("center" if sutun in (7, 8) else "right",
                                 "center", girinti=0 if sutun in (7, 8) else 1)

    # Durum ve oncelik sutunlarina rozet renkleri (kosullu bicimlendirme).
    u.durum_kosullu_bicim(ws, f"F{KONSOL_ILK_SATIR}:F2000")
    u.oncelik_kosullu_bicim(ws, f"I{KONSOL_ILK_SATIR}:I2000")

    # Baslik satiri ve ilk iki sutun sabit kalsin.
    ws.freeze_panes = f"C{KONSOL_ILK_SATIR}"
    ws.sheet_state = "veryHidden"
    return ws


# ==========================================================================
#  Degerlendirme
# ==========================================================================
def _degerlendirme(wb):
    ws = wb.create_sheet(SAYFA_DEGERLENDIRME)
    son_sutun = u.sayfa_hazirla(ws, DEG_SUTUNLAR, son_satir=52)
    u.masthead(ws, son_sutun, "Değerlendirme")

    u.bosluk(ws, 4, 8.0)
    u.bosluk(ws, 5, 40.0)       # dugmeler
    u.bosluk(ws, 6, 6.0)

    ws.row_dimensions[7].height = 22.0
    bant = u.birlestir(ws, 7, 2, 7, 6)
    bant.alignment = u.hiza("left", "center", girinti=1)
    wb.defined_names.add(_ad("dg_bant", SAYFA_DEGERLENDIRME, 2, 7))

    u.bosluk(ws, 8, 8.0)

    # --- Secili oneri (salt okunur) --------------------------------------
    u.bolum_basligi(ws, 9, 2, 6, "SEÇİLİ ÖNERİ",
                    "Konsolda bir satır seçip “Seçiliyi Değerlendir” düğmesine basın.")
    u.bosluk(ws, 11, 6.0)

    _okuma_alani(ws, wb, 12, 2, "ÖNERİ NO", "dg_oneri_no", 3, 3)
    _okuma_alani(ws, wb, 12, 5, "GÖNDERİM TARİHİ", "dg_tarih", 6, 6)
    _okuma_alani(ws, wb, 13, 2, "GÖNDEREN", "dg_gonderen", 3, 6)
    _okuma_alani(ws, wb, 14, 2, "ÖNERİ BAŞLIĞI", "dg_baslik", 3, 6)
    _okuma_alani(ws, wb, 15, 2, "MEVCUT DURUM", "dg_mevcut", 3, 6,
                 yukseklik=56.0, coklu=True)
    _okuma_alani(ws, wb, 16, 2, "ÇÖZÜM ÖNERİSİ", "dg_cozum", 3, 6,
                 yukseklik=56.0, coklu=True)
    _okuma_alani(ws, wb, 17, 2, "BEKLENEN FAYDA", "dg_fayda", 3, 6,
                 yukseklik=44.0, coklu=True)

    u.bosluk(ws, 18, 14.0)

    # --- Degerlendirme girisleri -----------------------------------------
    u.bolum_basligi(ws, 19, 2, 6, "DEĞERLENDİRME",
                    "Her kayıt yeni bir olay dosyası olarak eklenir; "
                    "önceki değerlendirmeler silinmez.")
    u.bosluk(ws, 21, 6.0)

    _giris_alani(ws, wb, 22, 2, "YENİ DURUM", "dg_yeni_durum", 3, 4,
                 liste="lst_durum", zorunlu=True)
    _giris_alani(ws, wb, 23, 2, "ETKİ  (1–5)", "dg_etki", 3, 3, liste="lst_puan")
    _giris_alani(ws, wb, 23, 5, "EFOR  (1–5)", "dg_efor", 6, 6, liste="lst_puan")

    # Oncelik sinifi hesaplanan bir degerdir; kullanici giremez.
    _okuma_alani(ws, wb, 24, 2, "ÖNCELİK SINIFI", "dg_oncelik", 3, 4)
    ws.cell(row=24, column=3).value = "— puanlanmadı —"

    _giris_alani(ws, wb, 25, 2, "YILLIK SAAT KAZANIMI", "dg_saat", 3, 4,
                 bicim="#,##0.0")
    _giris_alani(ws, wb, 26, 2, "YILLIK TL TASARRUFU", "dg_tl", 3, 4,
                 bicim="#,##0")
    _giris_alani(ws, wb, 27, 2, "KARAR NOTU", "dg_not", 3, 6,
                 yukseklik=48.0, coklu=True)

    u.ipucu_satiri(ws, 28, 3, 6,
                   "Kazanım rakamları yalnızca uygulamaya geçmiş önerilerde "
                   "panoya yansır. Reddedilen öneriler için gerekçe zorunludur.")
    u.bosluk(ws, 29, 12.0)

    # --- Gecmis -----------------------------------------------------------
    u.bolum_basligi(ws, 30, 2, 6, "DEĞERLENDİRME GEÇMİŞİ",
                    "En yeni kayıt üstte. Tam geçmiş ortak klasörde saklanır.")
    u.bosluk(ws, 32, 4.0)

    u.tablo_basligi(ws, DEG_GECMIS_BASLIK, 2, 6,
                    ["Tarih", "Değerlendiren", "Durum", "Puanlar", "Karar Notu"])
    u.tablo_govde_stili(ws, DEG_GECMIS_ILK, DEG_GECMIS_ILK + DEG_GECMIS_ADET - 1, 2, 6)
    u.durum_kosullu_bicim(
        ws, f"D{DEG_GECMIS_ILK}:D{DEG_GECMIS_ILK + DEG_GECMIS_ADET - 1}")

    ws.sheet_state = "veryHidden"
    return ws


def _okuma_alani(ws, wb, satir, etiket_sutun, etiket_metni, ad, c1, c2,
                 yukseklik=None, coklu=False):
    """Salt okunur bilgi alani: acik gri zemin, kilitli."""
    e = u.etiket(ws, satir, etiket_sutun, etiket_metni)
    if coklu:
        e.alignment = u.hiza("left", "top")
    if yukseklik:
        ws.row_dimensions[satir].height = yukseklik

    h = u.birlestir(ws, satir, c1, satir, c2)
    h.font = u.yazi(RENK["METIN_KOYU"], PT["GOVDE"], kalin=not coklu)
    h.alignment = u.hiza("left", "top" if coklu else "center",
                         kaydir=coklu, girinti=1)
    for c in range(c1, c2 + 1):
        hh = ws.cell(row=satir, column=c)
        hh.fill = u.dolgu(RENK["BEYAZ"])
        hh.border = u.Border(bottom=u.kenar(RENK["CIZGI_GRI"]))
    wb.defined_names.add(_ad(ad, SAYFA_DEGERLENDIRME, c1, satir))
    return h


def _giris_alani(ws, wb, satir, etiket_sutun, etiket_metni, ad, c1, c2,
                 liste=None, yukseklik=None, coklu=False, bicim=None,
                 zorunlu=False):
    """Kilitsiz giris alani (sayfa korumali oldugu icin yalnizca bunlar yazilabilir)."""
    e = u.etiket(ws, satir, etiket_sutun, etiket_metni, zorunlu=zorunlu)
    if coklu:
        e.alignment = u.hiza("left", "top")

    h = u.form_alani(ws, wb, satir, c1, c2, ad, SAYFA_DEGERLENDIRME,
                     yukseklik=yukseklik, coklu_satir=coklu)
    if bicim:
        for c in range(c1, c2 + 1):
            ws.cell(row=satir, column=c).number_format = bicim

    if liste:
        dv = DataValidation(type="list", formula1=f"={liste}", allow_blank=True,
                            showDropDown=False, showErrorMessage=True,
                            errorTitle="Geçersiz seçim",
                            error="Lütfen listeden bir değer seçin.")
        ws.add_data_validation(dv)
        dv.add(u.adres(satir, c1, satir, c2))
    return h


# ==========================================================================
#  Pano
# ==========================================================================
def _pano(wb):
    ws = wb.create_sheet(SAYFA_PANO)
    son_sutun = u.sayfa_hazirla(ws, PANO_SUTUNLAR, son_satir=52)
    u.masthead(ws, son_sutun, "Pano")

    u.bosluk(ws, 4, 8.0)
    u.bosluk(ws, 5, 40.0)       # dugmeler
    u.bosluk(ws, 6, 6.0)

    g = u.birlestir(ws, 7, 2, 7, 9)
    g.font = u.yazi(RENK["METIN_GRI"], 9)
    g.alignment = u.hiza("left", "center")
    ws.row_dimensions[7].height = 16.0
    wb.defined_names.add(_ad("pano_guncelleme", SAYFA_PANO, 2, 7))

    u.bosluk(ws, 8, 8.0)

    kartlar = [
        (9, [("Toplam öneri", "pano_toplam", "sisteme gelen tüm öneriler", "#,##0", None),
             ("Değerlendirme bekleyen", "pano_bekleyen", "Yeni + Değerlendirmede",
              "#,##0", RENK["KURUMSAL_MAVI"]),
             ("Uygulamaya geçmiş", "pano_uygulanan", "pilot, ölçüm ve standart",
              "#,##0", None)]),
        (13, [("Yıllık kazanılan saat", "pano_saat",
               "yalnızca uygulanmış önerilerden", "#,##0.0", None),
              ("Yıllık TL tasarrufu", "pano_tl",
               "yalnızca uygulanmış önerilerden", "#,##0 ₺", None),
              ("Hızlı kazanım", "pano_hizli", "yüksek etki · düşük efor",
               "#,##0", None)]),
        (17, [("Kabul oranı", "pano_kabul_orani", "sonuçlanan öneriler içinde",
               "0%", RENK["KURUMSAL_MAVI"]),
              ("Ortalama ilk yanıt", "pano_yanit_suresi",
               "gönderimden ilk değerlendirmeye (gün)", "#,##0.0",
               RENK["KURUMSAL_MAVI"]),
              ("Bu ay gelen", "pano_bu_ay", "içinde bulunduğumuz ay", "#,##0",
               RENK["KURUMSAL_MAVI"])]),
    ]
    for satir, grup in kartlar:
        for (c1, c2), (etiket, ad, alt, bicim, vurgu) in zip(PANO_KART_SUTUNLARI, grup):
            u.kpi_karti(ws, wb, satir, c1, c2, etiket, ad, SAYFA_PANO,
                        alt_metin=alt, sayi_bicimi=bicim, vurgu_hex=vurgu)
        u.bosluk(ws, satir + 3, 8.0)

    # --- Grafik basliklari ve alani --------------------------------------
    # Iki grafik: durum dagilimi solda, 12 aylik trend sagda ve daha genis.
    # Trend grafigi 12 kategori gosterdigi icin genislikten gercekten
    # yararlanir; dar bir kutuda ay etiketleri okunmaz hale gelirdi.
    u.bosluk(ws, 20, 14.0)
    for (c1, c2), metin in zip(PANO_GRAFIK_SUTUNLARI,
                               ["DURUM DAĞILIMI", "SON 12 AYIN GÖNDERİM TRENDİ"]):
        h = u.birlestir(ws, 21, c1, 21, c2)
        h.value = metin
        h.font = u.yazi(RENK["ANA_LACIVERT"], PT["BOLUM_BASLIK"], kalin=True,
                        aile=FONT_BASLIK_AILE)
        h.alignment = u.hiza("left", "bottom")
    ws.row_dimensions[21].height = 22.0
    u.bosluk(ws, 22, 4.0)

    for r in range(PANO_GRAFIK_UST, PANO_GRAFIK_ALT + 1):
        ws.row_dimensions[r].height = 16.5
    for (c1, c2) in PANO_GRAFIK_SUTUNLARI:
        u.blok_doldur(ws, PANO_GRAFIK_UST, c1, PANO_GRAFIK_ALT, c2, RENK["BEYAZ"])
        u.cerceve(ws, PANO_GRAFIK_UST, c1, PANO_GRAFIK_ALT, c2,
                  RENK["CIZGI_GRI"], golge_hex=RENK["GOLGE_GRI"])

    # --- Oncelik matrisi ---------------------------------------------------
    u.bosluk(ws, 39, 14.0)
    u.bolum_basligi(ws, 40, 2, 9, "ÖNCELİK MATRİSİ",
                    "Etki 3+ yüksek sayılır, efor 2− düşük sayılır.")
    u.bosluk(ws, 42, 4.0)
    _oncelik_matrisi(ws, wb)

    ws.sheet_state = "veryHidden"
    return ws


def _oncelik_matrisi(wb_ws, wb):
    """2×2 matris: Excel grafigi yerine renkli hucre bloklari.

    Dagilim scatter grafiginden cok daha okunakli: yonetici dort kutuda
    kac oneri oldugunu bir bakista gorur, nokta saymaz.
    """
    ws = wb_ws
    # Sutun eslesmeleri: satir etiketi B:C, dusuk efor E:F, yuksek efor H:I
    ETIKET = (2, 3)
    DUSUK = (5, 6)
    YUKSEK = (8, 9)

    ws.row_dimensions[43].height = 20.0
    for (c1, c2), metin in ((DUSUK, "DÜŞÜK EFOR  (1–2)"), (YUKSEK, "YÜKSEK EFOR  (3–5)")):
        h = u.birlestir(ws, 43, c1, 43, c2)
        h.value = metin
        h.font = u.yazi(RENK["METIN_GRI"], PT["ETIKET"], kalin=True)
        h.alignment = u.hiza("center", "center")

    satirlar = [
        (44, "YÜKSEK ETKİ  (3–5)",
         [(DUSUK, "Hızlı Kazanım", "pano_mtx_hizli", "önce bunlar"),
          (YUKSEK, "Büyük Proje", "pano_mtx_buyuk", "planlama ve kaynak ister")]),
        (48, "DÜŞÜK ETKİ  (1–2)",
         [(DUSUK, "Doldurma İşi", "pano_mtx_doldurma", "boş kapasiteyle"),
          (YUKSEK, "Değerlendirme Dışı", "pano_mtx_disi", "bu haliyle önerilmez")]),
    ]

    for ust, satir_etiketi, kutular in satirlar:
        ws.row_dimensions[ust].height = 34.0
        ws.row_dimensions[ust + 1].height = 18.0
        ws.row_dimensions[ust + 2].height = 14.0

        e = u.birlestir(ws, ust, ETIKET[0], ust + 2, ETIKET[1])
        e.value = satir_etiketi
        e.font = u.yazi(RENK["METIN_GRI"], PT["ETIKET"], kalin=True)
        e.alignment = u.hiza("right", "center")

        for (c1, c2), ad_metni, ad, ipucu in kutular:
            zemin, yazi_rengi = ONCELIK_RENK[ad_metni]
            u.blok_doldur(ws, ust, c1, ust + 2, c2, zemin)
            u.cerceve(ws, ust, c1, ust + 2, c2, RENK["CIZGI_GRI"],
                      golge_hex=RENK["GOLGE_GRI"])

            d = u.birlestir(ws, ust, c1, ust, c2)
            d.value = 0
            d.number_format = "#,##0"
            d.font = u.yazi(yazi_rengi, PT["KPI_RAKAM"], kalin=True,
                            aile=FONT_BASLIK_AILE)
            d.alignment = u.hiza("center", "center")
            wb.defined_names.add(_ad(ad, SAYFA_PANO, c1, ust))

            b = u.birlestir(ws, ust + 1, c1, ust + 1, c2)
            b.value = ad_metni
            b.font = u.yazi(yazi_rengi, 11, kalin=True)
            b.alignment = u.hiza("center", "center")

            i = u.birlestir(ws, ust + 2, c1, ust + 2, c2)
            i.value = ipucu
            i.font = u.yazi(RENK["METIN_GRI"], 9)
            i.alignment = u.hiza("center", "top")

        u.bosluk(ws, ust + 3, 8.0)


# ==========================================================================
#  Rapor
# ==========================================================================
def _rapor(wb):
    ws = wb.create_sheet(SAYFA_RAPOR)
    son_sutun = u.sayfa_hazirla(ws, RAPOR_SUTUNLAR, son_satir=30)
    u.masthead(ws, son_sutun, "Rapor")

    u.bosluk(ws, 4, 8.0)
    u.bosluk(ws, 5, 40.0)       # dugmeler
    u.bosluk(ws, 6, 10.0)

    u.bolum_basligi(ws, 7, 2, 5, "YÖNETİM RAPORU",
                    "Rapor Excel'in kendi AES şifrelemesiyle korunur; "
                    "parolasız açılamaz.")
    u.bosluk(ws, 9, 8.0)

    u.etiket(ws, 10, 2, "DÖNEM")
    u.form_alani(ws, wb, 10, 3, 4, "rpr_donem", SAYFA_RAPOR)
    ws.cell(row=10, column=3).value = "Tümü"
    dv = DataValidation(type="list", formula1="=lst_donem", allow_blank=False,
                        showDropDown=False, showErrorMessage=True,
                        errorTitle="Geçersiz seçim",
                        error="Lütfen listeden bir dönem seçin.")
    ws.add_data_validation(dv)
    dv.add(u.adres(10, 3, 10, 4))

    u.bosluk(ws, 11, 12.0)
    u.bilgi_kutusu(
        ws, 12, 2, 14, 5,
        "Parola üretim sırasında sorulur ve hiçbir yere kaydedilmez — bu "
        "dosyayı ele geçiren biri raporu açamaz. Parolayı kaybederseniz "
        "rapor kurtarılamaz; raporu yeniden üretmeniz gerekir.\n"
        "Raporda kazanım rakamlarına yalnızca uygulamaya geçmiş öneriler dahildir.",
    )
    for r in (12, 13, 14):
        ws.row_dimensions[r].height = 18.0

    u.bosluk(ws, 15, 12.0)
    s = u.birlestir(ws, 16, 2, 16, 5)
    s.font = u.yazi(RENK["METIN_GRI"], 9.5)
    s.alignment = u.hiza("left", "center")
    wb.defined_names.add(_ad("rpr_son", SAYFA_RAPOR, 2, 16))

    ws.sheet_state = "veryHidden"
    return ws


# ==========================================================================
#  Gizli calisma sayfalari
# ==========================================================================
VERI_BASLIKLARI = [
    "oneri_no", "tarih", "ad_soyad", "sicil_no",
    "mevcut_durum", "oneri_basligi", "cozum_onerisi", "beklenen_fayda",
    "durum", "etki", "efor", "oncelik", "yillik_saat", "yillik_tl",
    "ilk_olay", "son_olay", "degerlendiren", "karar_notu", "olay_sayisi",
    "gonderen_kullanici", "kaynak_dosya",
]


def _veri(wb):
    ws = wb.create_sheet(SAYFA_VERI)
    for i, ad in enumerate(VERI_BASLIKLARI, start=1):
        h = ws.cell(row=1, column=i, value=ad)
        h.font = u.yazi(RENK["BEYAZ"], 9, kalin=True)
        h.fill = u.dolgu(RENK["ANA_LACIVERT"])
        ws.column_dimensions[get_column_letter(i)].width = 20
    sayac_sutun = len(VERI_BASLIKLARI) + 2
    ws.cell(row=1, column=sayac_sutun, value="adet")
    ws.cell(row=2, column=sayac_sutun, value=0)
    wb.defined_names.add(_ad("veri_adet", SAYFA_VERI, sayac_sutun, 2))
    ws.sheet_state = "veryHidden"
    return ws


def _panoveri(wb):
    ws = wb.create_sheet(SAYFA_PANOVERI)
    for sutun, baslik in ((1, "durum"), (2, "adet"), (4, "ay"), (5, "adet")):
        h = ws.cell(row=1, column=sutun, value=baslik)
        h.font = u.yazi(RENK["METIN_GRI"], 9, kalin=True)
        ws.column_dimensions[get_column_letter(sutun)].width = 22

    # Etiket sutunlari METIN olarak biçimlendirilir. Aksi halde Excel "Mar 26"
    # gibi ay etiketlerini tarih sanip "26.Mar" diye yeniden yazar ve grafiğin
    # x ekseninde bazı aylar farklı görünür.
    for sutun in (1, 4):
        harf = get_column_letter(sutun)
        for satir in range(2, 16):
            ws[f"{harf}{satir}"].number_format = "@"

    ws.sheet_state = "veryHidden"
    return ws


# ==========================================================================
#  Dugmeler ve COM ek islemleri
# ==========================================================================
DUGMELER = [
    {"sayfa": SAYFA_GIRIS, "hucre": "C10", "metin": "Konsola Gir  →",
     "makro": "KonsolaGir", "birincil": True},

    {"sayfa": SAYFA_KONSOL, "hucre": "B5", "metin": "⟳  Önerileri Yenile",
     "makro": "OnerileriYenile", "birincil": True, "genislik": 160.0},
    {"sayfa": SAYFA_KONSOL, "hucre": "B5", "metin": "Seçiliyi Değerlendir",
     "makro": "SeciliyiDegerlendir", "birincil": False, "genislik": 168.0,
     "sol_kaydir": 170.0},
    {"sayfa": SAYFA_KONSOL, "hucre": "B5", "metin": "Pano",
     "makro": "PanoyaGit", "birincil": False, "genislik": 92.0,
     "sol_kaydir": 348.0},
    {"sayfa": SAYFA_KONSOL, "hucre": "B5", "metin": "Rapor",
     "makro": "RaporaGit", "birincil": False, "genislik": 92.0,
     "sol_kaydir": 450.0},
    {"sayfa": SAYFA_KONSOL, "hucre": "B5", "metin": "Çıkış",
     "makro": "Cikis", "birincil": False, "genislik": 80.0,
     "sol_kaydir": 552.0},

    {"sayfa": SAYFA_DEGERLENDIRME, "hucre": "B5",
     "metin": "✓  Değerlendirmeyi Kaydet", "makro": "DegerlendirmeKaydet",
     "birincil": True, "genislik": 196.0},
    {"sayfa": SAYFA_DEGERLENDIRME, "hucre": "B5", "metin": "←  Konsola Dön",
     "makro": "KonsolaDon", "birincil": False, "genislik": 140.0,
     "sol_kaydir": 206.0},

    {"sayfa": SAYFA_PANO, "hucre": "B5", "metin": "⟳  Panoyu Yenile",
     "makro": "PanoyuYenile", "birincil": True, "genislik": 156.0},
    {"sayfa": SAYFA_PANO, "hucre": "B5", "metin": "←  Konsola Dön",
     "makro": "KonsolaDon", "birincil": False, "genislik": 140.0,
     "sol_kaydir": 166.0},
    {"sayfa": SAYFA_PANO, "hucre": "B5", "metin": "Rapor",
     "makro": "RaporaGit", "birincil": False, "genislik": 92.0,
     "sol_kaydir": 316.0},

    {"sayfa": SAYFA_RAPOR, "hucre": "B5", "metin": "🔒  Parolalı Rapor Üret",
     "makro": "RaporUret", "birincil": True, "genislik": 196.0},
    {"sayfa": SAYFA_RAPOR, "hucre": "B5", "metin": "←  Konsola Dön",
     "makro": "KonsolaDon", "birincil": False, "genislik": 140.0,
     "sol_kaydir": 206.0},
]


def com_ek_islem(wb):
    """Pano grafiklerini kurar. openpyxl canli grafik baglayamaz; bu is
    gercek Excel uzerinden yapilir."""
    from tasarim import rgb_long

    XL_BAR = 57            # xlBarClustered
    XL_COLUMN = 51         # xlColumnClustered
    XL_LINE_MARKERS = 65   # xlLineMarkers
    XL_CATEGORY, XL_VALUE = 1, 2

    pano = wb.Worksheets(SAYFA_PANO)
    eski = pano.Visible
    pano.Visible = -1

    # Durum grafiginin cubuklari konsoldaki rozetlerle AYNI renkleri kullanir;
    # boylece iki ekran ayni renk dilini konusur.
    durum_renkleri = [DURUM_RENK[d][1] for d in
                      ["Yeni", "Değerlendirmede", "Planlandı", "Pilot Uygulamada",
                       "Ölçümleniyor", "Standartlaştırıldı", "Beklemede",
                       "Reddedildi"]]

    # Etiket ve deger araliklari AYRI AYRI verilir; seri elle kurulur.
    # SetSourceData'nin otomatik tahmini guvenilir degil: bir aralikta ilk veri
    # satirini baslik sanip kategoriyi dusuruyor, digerinde etiket sutununu
    # ikinci bir seri sanip kategori adlarini 1, 2, 3'e ceviriyordu.
    tanimlar = [
        (PANO_GRAFIK_SUTUNLARI[0], "$A$2:$A$9", "$B$2:$B$9", XL_BAR,
         RENK["ANA_LACIVERT"], True, durum_renkleri),
        (PANO_GRAFIK_SUTUNLARI[1], "$D$2:$D$13", "$E$2:$E$13", XL_LINE_MARKERS,
         RENK["ANA_LACIVERT"], False, None),
    ]

    pv = wb.Worksheets(SAYFA_PANOVERI)

    for (c1, c2), etiket_aralik, deger_aralik, tur, renk, ters, nokta_renkleri \
            in tanimlar:
        sol = pano.Cells(PANO_GRAFIK_UST, c1).Left + 4
        ust = pano.Cells(PANO_GRAFIK_UST, c1).Top + 4
        sag = pano.Cells(PANO_GRAFIK_ALT, c2).Left + pano.Cells(PANO_GRAFIK_ALT, c2).Width
        alt = pano.Cells(PANO_GRAFIK_ALT, c2).Top + pano.Cells(PANO_GRAFIK_ALT, c2).Height

        co = pano.ChartObjects().Add(sol, ust, sag - sol - 8, alt - ust - 8)
        co.Name = f"grafik_{c1}"
        ch = co.Chart
        ch.ChartType = tur

        while ch.SeriesCollection().Count > 0:
            ch.SeriesCollection(1).Delete()
        yeni = ch.SeriesCollection().NewSeries()
        yeni.Values = pv.Range(deger_aralik)
        yeni.XValues = pv.Range(etiket_aralik)

        try:
            ch.ChartArea.Format.Fill.ForeColor.RGB = rgb_long(RENK["BEYAZ"])
            ch.ChartArea.Format.Line.Visible = 0
            ch.ChartArea.Font.Name = "Segoe UI"
            ch.ChartArea.Font.Size = 9
            ch.ChartArea.Font.Color = rgb_long(RENK["METIN_GRI"])
            ch.PlotArea.Format.Fill.Visible = 0
        except Exception:
            pass

        try:
            seri = ch.SeriesCollection(1)
            seri.Format.Fill.ForeColor.RGB = rgb_long(renk)
            seri.Format.Line.ForeColor.RGB = rgb_long(renk)
            seri.Format.Line.Weight = 2.0
            seri.HasDataLabels = True
            seri.DataLabels().Font.Size = 9
            seri.DataLabels().Font.Name = "Segoe UI"
            seri.DataLabels().Font.Color = rgb_long(RENK["METIN_KOYU"])

            if nokta_renkleri:
                for i, hex_kod in enumerate(nokta_renkleri, start=1):
                    seri.Points(i).Format.Fill.ForeColor.RGB = rgb_long(hex_kod)
        except Exception:
            pass

        try:
            kategori = ch.Axes(XL_CATEGORY)
            kategori.TickLabels.Font.Size = 9
            kategori.Format.Line.ForeColor.RGB = rgb_long(RENK["CIZGI_GRI"])
            if ters:
                kategori.ReversePlotOrder = True

            # Sayilar tam sayidir; eksen 0,5 gibi ara degerler gostermemeli.
            deger = ch.Axes(XL_VALUE)
            deger.MinimumScale = 0
            deger.MajorUnit = 1
            deger.TickLabels.NumberFormat = "0"
            deger.TickLabels.Font.Size = 9
            deger.MajorGridlines.Format.Line.ForeColor.RGB = rgb_long(RENK["CIZGI_GRI"])
        except Exception:
            pass

        # Baslik ve gosterge EN SONDA kapatilir: ChartType ve SetSourceData
        # cagrilari bunlari yeniden acabiliyor.
        try:
            ch.HasLegend = False
            ch.HasTitle = False
        except Exception:
            pass

    pano.Visible = eski


# ==========================================================================
def kitap_uret(hedef_yol, listeler):
    wb = Workbook()
    wb.remove(wb.active)

    _giris(wb)
    _konsol(wb)
    _degerlendirme(wb)
    _pano(wb)
    _rapor(wb)
    _veri(wb)
    _panoveri(wb)

    tum_listeler = dict(listeler)
    tum_listeler["donem"] = ["Tümü", "Bu Yıl", "Bu Ay"]
    u.listeler_sayfasi_kur(wb, tum_listeler)

    wb.save(hedef_yol)
    return hedef_yol
