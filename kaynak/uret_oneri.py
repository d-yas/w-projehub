# -*- coding: utf-8 -*-
"""KaizenOneri.xlsm -- personel gonderim kitabinin sayfa kurulumu.

Bu betik yalnizca gorsel/yapisal iskeleti kurar (.xlsx taslak). VBA modulleri,
dugme sekilleri ve korumalar com_kurulum.py asamasinda eklenir.
"""

from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

import uret_ortak as u
from tasarim import FONT_BASLIK_AILE, PT, RENK

SAYFA_GIRIS = "Giriş"
SAYFA_FORM = "Öneri Formu"

# A: kenar boslugu, B: etiket, C-F: giris alani, G: kenar boslugu
SUTUNLAR = [2.2, 24.0, 20.0, 20.0, 20.0, 14.0, 2.2]
SON_SUTUN = len(SUTUNLAR)
C_ETIKET = 2
C_ALAN_BAS = 3
C_ALAN_SON = 6

# Dugmelerin uzerine oturacagi hucreler (COM asamasi bu adresleri kullanir)
DUGME_YERLERI = {
    "giris": "C10",
    "gonder": "C32",
    "temizle": "E32",
    "cikis": "F5",
}


def _giris_sayfasi(wb):
    ws = wb.create_sheet(SAYFA_GIRIS)
    u.sayfa_hazirla(ws, SUTUNLAR, son_satir=40)
    u.masthead(ws, SON_SUTUN, "Giriş")

    u.bosluk(ws, 4, 18.0)

    baslik = u.birlestir(ws, 5, C_ETIKET, 5, C_ALAN_SON)
    baslik.value = "Proje Öneri Formu"
    baslik.font = u.yazi(RENK["ANA_LACIVERT"], 20, kalin=True, aile=FONT_BASLIK_AILE)
    baslik.alignment = u.hiza("left", "center")
    ws.row_dimensions[5].height = 30.0

    alt = u.birlestir(ws, 6, C_ETIKET, 6, C_ALAN_SON)
    alt.value = "İşinizi kolaylaştıracak her fikir bir Kaizen'dir."
    alt.font = u.yazi(RENK["METIN_GRI"], 11)
    alt.alignment = u.hiza("left", "center")
    ws.row_dimensions[6].height = 20.0

    u.bosluk(ws, 7, 14.0)

    u.bilgi_kutusu(
        ws, 8, C_ETIKET, 9, C_ALAN_SON,
        "Öneriniz Kaizen ekibi tarafından PDCA döngüsüyle değerlendirilir. "
        "Gönderim sonrasında size bir öneri numarası verilir; önerinizin "
        "durumunu bu numarayla sorabilirsiniz.",
    )
    ws.row_dimensions[8].height = 20.0
    ws.row_dimensions[9].height = 20.0

    u.bosluk(ws, 10, 44.0)      # "Sisteme Gir" dugmesi buraya oturur
    u.bosluk(ws, 11, 12.0)

    not_ = u.birlestir(ws, 12, C_ETIKET, 12, C_ALAN_SON)
    not_.value = ("Dosya açıldığında üstte bir güvenlik uyarısı çıkarsa "
                  "“İçeriği Etkinleştir” düğmesine basın.")
    not_.font = u.yazi(RENK["METIN_GRI"], 9, italik=True)
    not_.alignment = u.hiza("left", "center")

    # Kaizen dongusu kisa anlatimi -- ekrani bos birakmamak icin degil,
    # kullaniciya sistemin ne yaptigini anlatmak icin.
    u.bosluk(ws, 14, 10.0)
    u.bolum_basligi(ws, 15, C_ETIKET, C_ALAN_SON, "Öneriniz nasıl ilerler?")
    adimlar = [
        ("1  Gönderirsiniz", "Form doldurulur, öneri numarası verilir."),
        ("2  Değerlendirilir", "Kaizen ekibi etki ve efor puanlar, önceliklendirir."),
        ("3  Uygulanır", "Kabul edilen öneri pilot uygulamayla denenir."),
        ("4  Standartlaşır", "Sonuç ölçülür, işe yarayan yaygınlaştırılır."),
    ]
    r = 17
    for ad, aciklama in adimlar:
        h = ws.cell(row=r, column=C_ETIKET)
        h.value = ad
        h.font = u.yazi(RENK["KURUMSAL_MAVI"], 10, kalin=True)
        h.alignment = u.hiza("left", "center")
        a = u.birlestir(ws, r, C_ALAN_BAS, r, C_ALAN_SON)
        a.value = aciklama
        a.font = u.yazi(RENK["METIN_GRI"], 10)
        a.alignment = u.hiza("left", "center")
        ws.row_dimensions[r].height = 18.0
        r += 1

    ws.sheet_state = "visible"
    return ws


def _form_sayfasi(wb):
    ws = wb.create_sheet(SAYFA_FORM)
    u.sayfa_hazirla(ws, SUTUNLAR, son_satir=44)
    u.masthead(ws, SON_SUTUN, "Öneri Formu")

    u.bosluk(ws, 4, 10.0)

    # Durum bandi: gonderim sonucunu formu terk etmeden gosterir.
    ws.row_dimensions[5].height = 24.0
    bant = u.birlestir(ws, 5, C_ETIKET, 5, C_ALAN_SON - 1)
    bant.alignment = u.hiza("left", "center", girinti=1)
    wb.defined_names.add(_ad("frm_bant", SAYFA_FORM, C_ETIKET, 5))

    u.bosluk(ws, 6, 10.0)

    # --- 1. Kimlik --------------------------------------------------------
    u.bolum_basligi(ws, 7, C_ETIKET, C_ALAN_SON, "1  ·  KİMLİK BİLGİLERİ",
                    "Önerinizle ilgili sorumuz olursa size ulaşabilmemiz için.")
    u.bosluk(ws, 9, 10.0)

    _alan(ws, wb, 10, "Ad Soyad", "frm_ad_soyad", zorunlu=True)
    u.bosluk(ws, 11, 5.0)
    _alan(ws, wb, 12, "Sicil No", "frm_sicil_no", zorunlu=True)

    u.bosluk(ws, 13, 18.0)

    # --- 2. Mevcut durum --------------------------------------------------
    u.bolum_basligi(ws, 14, C_ETIKET, C_ALAN_SON, "2  ·  ŞU ANKİ DURUM",
                    "Neyin iyileştirilmesi gerektiğini anlatın.")
    u.bosluk(ws, 16, 10.0)

    _alan(ws, wb, 17, "Mevcut Durum", "frm_mevcut", zorunlu=True,
          yukseklik=66.0, coklu=True)
    u.ipucu_satiri(ws, 18, C_ALAN_BAS, C_ALAN_SON,
                   "Alt satıra geçmek için Alt + Enter kullanın.")

    u.bosluk(ws, 19, 18.0)

    # --- 3. Oneri ---------------------------------------------------------
    u.bolum_basligi(ws, 20, C_ETIKET, C_ALAN_SON, "3  ·  ÖNERİNİZ",
                    "Nasıl daha iyi olabileceğini ve ne kazandıracağını yazın.")
    u.bosluk(ws, 22, 10.0)

    _alan(ws, wb, 23, "Öneri Başlığı", "frm_baslik", zorunlu=True)
    u.bosluk(ws, 24, 5.0)
    _alan(ws, wb, 25, "Çözüm Öneriniz", "frm_cozum", zorunlu=True,
          yukseklik=66.0, coklu=True)
    u.bosluk(ws, 26, 5.0)
    _alan(ws, wb, 27, "Beklenen Fayda", "frm_fayda", yukseklik=52.0, coklu=True)
    u.ipucu_satiri(ws, 28, C_ALAN_BAS, C_ALAN_SON,
                   "Örnek: “İşlem başına 4 dakika kazanç, günde ~30 işlem.”")

    # --- Dugmeler ---------------------------------------------------------
    u.bosluk(ws, 29, 14.0)
    u.bosluk(ws, 30, 1.0)
    u.blok_doldur(ws, 30, C_ETIKET, 30, C_ALAN_SON, RENK["CIZGI_GRI"])
    u.bosluk(ws, 31, 12.0)
    u.bosluk(ws, 32, 40.0)      # Gonder / Temizle dugmeleri
    u.bosluk(ws, 33, 14.0)

    u.bilgi_kutusu(
        ws, 34, C_ETIKET, 35, C_ALAN_SON,
        "Gönderdiğiniz kayıt ortak klasöre kendi dosyası olarak yazılır; "
        "başka kimsenin gönderimiyle çakışmaz. Kim, ne zaman, ne gönderdi "
        "bilgisi izlenebilirlik için kayda eklenir.",
        zemin=RENK["BUZ_ZEMIN"], yazi_hex=RENK["METIN_GRI"],
    )
    ws.row_dimensions[34].height = 18.0
    ws.row_dimensions[35].height = 18.0

    ws.sheet_state = "veryHidden"
    return ws


def _ad(ad, sayfa, sutun, satir):
    from openpyxl.utils import get_column_letter
    from openpyxl.workbook.defined_name import DefinedName
    return DefinedName(ad, attr_text=f"'{sayfa}'!${get_column_letter(sutun)}${satir}")


def _alan(ws, wb, satir, etiket_metni, ad, zorunlu=False, liste=None,
          yukseklik=None, coklu=False):
    """Bir etiket + giris alani satiri kurar."""
    e = u.etiket(ws, satir, C_ETIKET, u.buyuk(etiket_metni), zorunlu=zorunlu)
    if coklu:
        e.alignment = u.hiza("left", "top")
    if yukseklik is None:
        ws.row_dimensions[satir].height = 22.0      # rahat yazma yüksekliği

    u.form_alani(ws, wb, satir, C_ALAN_BAS, C_ALAN_SON, ad, SAYFA_FORM,
                 yukseklik=yukseklik, coklu_satir=coklu)

    if liste:
        dv = DataValidation(type="list", formula1=f"={liste}", allow_blank=True,
                            showDropDown=False, showErrorMessage=True,
                            errorTitle="Geçersiz seçim",
                            error="Lütfen listeden bir değer seçin.")
        ws.add_data_validation(dv)
        dv.add(u.adres(satir, C_ALAN_BAS, satir, C_ALAN_SON))


def kitap_uret(hedef_yol, listeler):
    wb = Workbook()
    wb.remove(wb.active)

    _giris_sayfasi(wb)
    _form_sayfasi(wb)
    u.listeler_sayfasi_kur(wb, listeler)

    wb.save(hedef_yol)
    return hedef_yol
