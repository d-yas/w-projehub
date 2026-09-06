# -*- coding: utf-8 -*-
"""openpyxl tasarim yardimcilari.

Bu modul "Excel'de nasil iyi gorunur" sorusunun tek cevap yeridir: masthead,
kart, bolum basligi, form alani, tablo basligi ve rozet desenleri burada
tanimlanir. Sayfa uretici betikler (uret_oneri, uret_yonetim) yalnizca bu
desenleri cagirir; boylece iki kitap gorsel olarak ayni dili konusur.
"""

from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from tasarim import (DURUM_RENK, FONT_AILE, FONT_BASLIK_AILE, OLCU,
                     ONCELIK_RENK, PT, RENK, URUN_ADI)


# --------------------------------------------------------------------------
# Temel stil kuruculari
# --------------------------------------------------------------------------
def argb(hex_kod: str) -> str:
    return "FF" + hex_kod.lstrip("#").upper()


def dolgu(hex_kod: str) -> PatternFill:
    return PatternFill(start_color=argb(hex_kod), end_color=argb(hex_kod), fill_type="solid")


def yazi(hex_kod=RENK["METIN_KOYU"], boyut=PT["GOVDE"], kalin=False, aile=None, italik=False):
    return Font(name=aile or FONT_AILE, size=boyut, bold=kalin,
                italic=italik, color=argb(hex_kod))


def kenar(hex_kod=RENK["CIZGI_GRI"], stil="thin") -> Side:
    return Side(style=stil, color=argb(hex_kod))


def hiza(yatay="left", dikey="center", kaydir=False, girinti=0) -> Alignment:
    return Alignment(horizontal=yatay, vertical=dikey, wrap_text=kaydir, indent=girinti)


def buyuk(metin: str) -> str:
    """Türkçe'ye uygun büyük harf.

    Python'un upper() metodu 'i' harfini 'I' yapar; Türkçe'de doğrusu 'İ'dir.
    Aksi halde ekranlarda "SICIL NO", "ÖNERI FORMU", "DEĞERLENDIRME" gibi
    yanlış yazımlar çıkar. Noktasız 'ı' da doğru karşılığına ('I') çevrilir.
    """
    return metin.replace("i", "İ").replace("ı", "I").upper()


# --------------------------------------------------------------------------
# Aralik islemleri
# --------------------------------------------------------------------------
def aralik(ws, r1, c1, r2, c2):
    """Satir/sutun numaralariyla hucre listesi dondurur."""
    return [ws.cell(row=r, column=c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)]


def adres(r1, c1, r2=None, c2=None) -> str:
    if r2 is None:
        return f"{get_column_letter(c1)}{r1}"
    return f"{get_column_letter(c1)}{r1}:{get_column_letter(c2)}{r2}"


def blok_doldur(ws, r1, c1, r2, c2, hex_kod):
    d = dolgu(hex_kod)
    for h in aralik(ws, r1, c1, r2, c2):
        h.fill = d


def cerceve(ws, r1, c1, r2, c2, hex_kod=RENK["CIZGI_GRI"],
            golge_hex=None, stil="thin"):
    """Blogun disina cerceve cizer. golge_hex verilirse sag ve alt kenar
    daha koyu cizilir; Excel'de golge olmadigi icin derinlik hissi boyle verilir."""
    ince = kenar(hex_kod, stil)
    koyu = kenar(golge_hex, "medium") if golge_hex else ince
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            h = ws.cell(row=r, column=c)
            h.border = Border(
                top=ince if r == r1 else None,
                bottom=(koyu if r == r2 else None),
                left=ince if c == c1 else None,
                right=(koyu if c == c2 else None),
            )


def birlestir(ws, r1, c1, r2, c2):
    """Hucreleri birlestirir.

    Tek hucrelik "birlestirme" (C24:C24) Excel'de gecerli ama bozuk bir
    durumdur: hucre birlesik sayilir ve ClearContents gibi islemler
    "birleştirilmiş bir hücrede bunu yapamayız" hatasi verir. Bu yuzden
    1×1 aralikta birlestirme yapilmaz.
    """
    if r1 != r2 or c1 != c2:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    return ws.cell(row=r1, column=c1)


# --------------------------------------------------------------------------
# Sayfa iskeleti
# --------------------------------------------------------------------------
def sayfa_hazirla(ws, sutun_genislikleri, zemin_hex=RENK["BUZ_ZEMIN"],
                  son_satir=90, baslik_gizle=True):
    """Kilavuz cizgilerini kapatir, zemini boyar, sutun genisliklerini kurar.

    Kilavuz cizgileri kapali + tek renk zemin, Excel sayfasini hucre izgarasi
    gibi degil bir uygulama ekrani gibi gosteren tek en etkili ayardir.
    """
    ws.sheet_view.showGridLines = False
    if baslik_gizle:
        ws.sheet_view.showRowColHeaders = False
    ws.sheet_view.zoomScale = 100

    for i, genislik in enumerate(sutun_genislikleri, start=1):
        ws.column_dimensions[get_column_letter(i)].width = genislik

    son_sutun = len(sutun_genislikleri)
    blok_doldur(ws, 1, 1, son_satir, son_sutun, zemin_hex)
    for r in range(1, son_satir + 1):
        if r not in ws.row_dimensions or ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = OLCU["SATIR_NORMAL"]
    return son_sutun


def masthead(ws, son_sutun, sayfa_adi, ust_metin=URUN_ADI):
    """Her sayfanin ustundeki iki satirlik lacivert bant.

    Solda sistemin adi, sagda bulundugunuz ekran. Kullanicinin nerede
    oldugunu bir bakista anlamasini saglar.
    """
    ws.row_dimensions[1].height = OLCU["MASTHEAD_YUKSEKLIK"]
    ws.row_dimensions[2].height = OLCU["MASTHEAD_YUKSEKLIK"]
    blok_doldur(ws, 1, 1, 2, son_sutun, RENK["ANA_LACIVERT"])

    # Bant ikiye bolunur ama YARI YARIYA degil: sagdaki yalnizca kisa bir sayfa
    # adi tasir (9-10 karakter), soldaki ise urun adinin tamamini. Yariya
    # bolunseydi uzun bir urun adi ("XJ Birimi | Proje Öneri Formu") kirpilirdi.
    sag_bas = max(3, son_sutun - 2)
    sol = birlestir(ws, 1, 2, 2, sag_bas - 1)
    sol.value = ust_metin
    sol.font = yazi(RENK["BEYAZ"], PT["SAYFA_BASLIK"], kalin=True, aile=FONT_BASLIK_AILE)
    sol.alignment = hiza("left", "center")

    sag = birlestir(ws, 1, sag_bas, 2, son_sutun - 1)
    sag.value = buyuk(sayfa_adi)
    sag.font = yazi(RENK["SOLUK_MAVI"], PT["MASTHEAD_ALT"], kalin=True)
    sag.alignment = hiza("right", "center")

    # Bandin altina ince bir vurgu cizgisi
    ws.row_dimensions[3].height = 3.0
    blok_doldur(ws, 3, 1, 3, son_sutun, RENK["KURUMSAL_MAVI"])


def bolum_basligi(ws, satir, c1, c2, metin, alt_metin=""):
    """Numarali bolum basligi: '1 · KİMLİK BİLGİLERİ'."""
    ws.row_dimensions[satir].height = 22.0
    h = birlestir(ws, satir, c1, satir, c2)
    h.value = metin
    h.font = yazi(RENK["ANA_LACIVERT"], PT["BOLUM_BASLIK"], kalin=True, aile=FONT_BASLIK_AILE)
    h.alignment = hiza("left", "bottom")
    if alt_metin:
        ws.row_dimensions[satir + 1].height = 14.0
        a = birlestir(ws, satir + 1, c1, satir + 1, c2)
        a.value = alt_metin
        a.font = yazi(RENK["METIN_GRI"], PT["ETIKET"])
        a.alignment = hiza("left", "top")


def etiket(ws, satir, sutun, metin, zorunlu=False):
    h = ws.cell(row=satir, column=sutun)
    h.value = (metin + " *") if zorunlu else metin
    h.font = yazi(RENK["METIN_GRI"], PT["ETIKET"], kalin=True)
    h.alignment = hiza("left", "center")
    return h


def form_alani(ws, wb, satir, c1, c2, ad, sayfa_adi, yukseklik=None,
               coklu_satir=False, ipucu=""):
    """Beyaz zeminli, alt kenarligi olan, KILITSIZ bir giris alani.

    Sayfa korumali oldugu icin yalnizca bu hucreler yazilabilir; kullanici
    tasarimi yanlislikla bozamaz. Alan adlandirilmis aralik olarak kaydedilir,
    VBA hucre adresi degil bu adi kullanir.
    """
    if yukseklik:
        ws.row_dimensions[satir].height = yukseklik

    h = birlestir(ws, satir, c1, satir, c2)
    h.fill = dolgu(RENK["BEYAZ"])
    h.font = yazi(RENK["METIN_KOYU"], PT["GOVDE"])
    h.alignment = hiza("left", "top" if coklu_satir else "center",
                       kaydir=coklu_satir, girinti=1)
    alt = kenar(RENK["ALAN_CIZGI"])
    yan = kenar(RENK["CIZGI_GRI"])
    for c in range(c1, c2 + 1):
        hh = ws.cell(row=satir, column=c)
        hh.border = Border(top=yan, bottom=alt,
                           left=yan if c == c1 else None,
                           right=yan if c == c2 else None)
        hh.fill = dolgu(RENK["BEYAZ"])
        hh.protection = Protection(locked=False)

    if ipucu:
        h.comment = None  # ipucu satir altinda gosterilir, acilir not kullanilmaz

    wb.defined_names.add(
        DefinedName(ad, attr_text=f"'{sayfa_adi}'!${get_column_letter(c1)}${satir}")
    )
    return h


def ipucu_satiri(ws, satir, c1, c2, metin):
    ws.row_dimensions[satir].height = 13.0
    h = birlestir(ws, satir, c1, satir, c2)
    h.value = metin
    h.font = yazi(RENK["METIN_GRI"], 8.5, italik=True)
    h.alignment = hiza("left", "center", girinti=1)
    return h


def bosluk(ws, satir, yukseklik=OLCU["SATIR_BOSLUK"]):
    ws.row_dimensions[satir].height = yukseklik


# --------------------------------------------------------------------------
# KPI karti
# --------------------------------------------------------------------------
def kpi_karti(ws, wb, r1, c1, c2, etiket_metni, ad, sayfa_adi,
              alt_metin="", sayi_bicimi="#,##0", vurgu_hex=None):
    """Uc satirlik kart: ustte kucuk gri etiket, ortada buyuk rakam, altta not.

    r1 karttin ilk satiridir; kart r1..r1+2 arasini kaplar.
    Deger hucresi adlandirilmis aralik olur, modPano oraya yazar.
    """
    vurgu = vurgu_hex or RENK["ANA_LACIVERT"]

    ws.row_dimensions[r1].height = 16.0
    ws.row_dimensions[r1 + 1].height = 32.0
    ws.row_dimensions[r1 + 2].height = 14.0

    blok_doldur(ws, r1, c1, r1 + 2, c2, RENK["BEYAZ"])
    cerceve(ws, r1, c1, r1 + 2, c2, RENK["CIZGI_GRI"], golge_hex=RENK["GOLGE_GRI"])

    e = birlestir(ws, r1, c1, r1, c2)
    e.value = buyuk(etiket_metni)
    e.font = yazi(RENK["METIN_GRI"], PT["KPI_ETIKET"], kalin=True)
    e.alignment = hiza("left", "center", girinti=1)

    d = birlestir(ws, r1 + 1, c1, r1 + 1, c2)
    d.value = 0
    d.number_format = sayi_bicimi
    d.font = yazi(vurgu, PT["KPI_RAKAM"], kalin=True, aile=FONT_BASLIK_AILE)
    d.alignment = hiza("left", "center", girinti=1)

    a = birlestir(ws, r1 + 2, c1, r1 + 2, c2)
    a.value = alt_metin
    a.font = yazi(RENK["METIN_GRI"], PT["KPI_ALT"])
    a.alignment = hiza("left", "top", girinti=1)

    wb.defined_names.add(
        DefinedName(ad, attr_text=f"'{sayfa_adi}'!${get_column_letter(c1)}${r1 + 1}")
    )
    return d


# --------------------------------------------------------------------------
# Tablo
# --------------------------------------------------------------------------
def tablo_basligi(ws, satir, c1, c2, basliklar):
    ws.row_dimensions[satir].height = 26.0
    blok_doldur(ws, satir, c1, satir, c2, RENK["ANA_LACIVERT"])
    for i, metin in enumerate(basliklar):
        h = ws.cell(row=satir, column=c1 + i)
        h.value = metin
        h.font = yazi(RENK["BEYAZ"], PT["TABLO_BASLIK"], kalin=True)
        h.alignment = hiza("center", "center", kaydir=True)
        h.border = Border(right=kenar(RENK["KURUMSAL_MAVI"]))


def tablo_govde_stili(ws, r1, r2, c1, c2, serit=True):
    """Tablo govdesine ince cizgi ve tek/cift satir seridi uygular."""
    ince = kenar(RENK["CIZGI_GRI"])
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].height = OLCU["SATIR_NORMAL"]
        zemin = RENK["BEYAZ"]
        if serit and (r - r1) % 2 == 1:
            zemin = RENK["BUZ_ZEMIN"]
        for c in range(c1, c2 + 1):
            h = ws.cell(row=r, column=c)
            h.fill = dolgu(zemin)
            h.border = Border(bottom=ince, right=ince)
            h.font = yazi(RENK["METIN_KOYU"], PT["GOVDE"])
            h.alignment = hiza("left", "center", girinti=1)


def bilgi_kutusu(ws, r1, c1, r2, c2, metin, zemin=RENK["ACIK_MAVI"],
                 yazi_hex=RENK["ANA_LACIVERT"]):
    blok_doldur(ws, r1, c1, r2, c2, zemin)
    cerceve(ws, r1, c1, r2, c2, RENK["CIZGI_GRI"])
    h = birlestir(ws, r1, c1, r2, c2)
    h.value = metin
    h.font = yazi(yazi_hex, PT["GOVDE"])
    h.alignment = hiza("left", "center", kaydir=True, girinti=1)
    return h


def dugme_yeri(ws, satir, sutun, yukseklik=34.0):
    """Dugme sekli icin yer acar. Sekiller COM asamasinda eklenir; openpyxl
    yuvarlatilmis dikdortgen + OnAction bagi kuramaz."""
    ws.row_dimensions[satir].height = yukseklik
    return adres(satir, sutun)


# --------------------------------------------------------------------------
# Kosullu bicimlendirme -- durum ve oncelik rozetleri
# --------------------------------------------------------------------------
def durum_kosullu_bicim(ws, hucre_araligi):
    """Durum sutununa her durum icin bir kural ekler.

    Kosullu bicimlendirme tercih edilir cunku VBA tabloyu yeniden kurdugunda
    renkleri ayrica boyamasi gerekmez; kural sayfada kalicidir.
    """
    from openpyxl.formatting.rule import CellIsRule

    for durum, (zemin, yazi_renk) in DURUM_RENK.items():
        ws.conditional_formatting.add(
            hucre_araligi,
            CellIsRule(operator="equal", formula=[f'"{durum}"'],
                       fill=dolgu(zemin), font=yazi(yazi_renk, PT["GOVDE"], kalin=True)),
        )


def oncelik_kosullu_bicim(ws, hucre_araligi):
    from openpyxl.formatting.rule import CellIsRule

    for sinif, (zemin, yazi_renk) in ONCELIK_RENK.items():
        ws.conditional_formatting.add(
            hucre_araligi,
            CellIsRule(operator="equal", formula=[f'"{sinif}"'],
                       fill=dolgu(zemin), font=yazi(yazi_renk, PT["GOVDE"], kalin=True)),
        )


# --------------------------------------------------------------------------
# Listeler sayfasi (veri dogrulama kaynagi)
# --------------------------------------------------------------------------
def listeler_sayfasi_kur(wb, listeler: dict):
    """Gizli 'Listeler' sayfasina her listeyi bir sutun olarak yazar ve
    her biri icin adlandirilmis aralik olusturur (lst_durum gibi)."""
    ws = wb.create_sheet("Listeler")
    for i, (ad, degerler) in enumerate(listeler.items(), start=1):
        sutun = get_column_letter(i)
        ws.cell(row=1, column=i, value=ad).font = yazi(RENK["METIN_GRI"], 9, kalin=True)
        for j, deger in enumerate(degerler, start=2):
            ws.cell(row=j, column=i, value=deger)
        ws.column_dimensions[sutun].width = 24
        wb.defined_names.add(
            DefinedName(f"lst_{ad}",
                        attr_text=f"Listeler!${sutun}$2:${sutun}${len(degerler) + 1}")
        )
    ws.sheet_state = "hidden"
    return ws
