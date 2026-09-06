# -*- coding: utf-8 -*-
"""openpyxl tasarim yardimcilari -- gorsel dilin tek tanim yeri.

Bu modul "Excel'de nasil iyi gorunur" sorusunun tek cevap yeridir: masthead,
kart, bolum basligi, form alani, tablo ve rozet desenleri burada tanimlanir.
Sayfa uretici betikler (uret_oneri, uret_yonetim) yalnizca bu desenleri
cagirir; boylece iki kitap gorsel olarak ayni dili konusur.

Gorsel dilin uc kurali:
  1. ZEMIN acik, ICERIK beyaz. Sayfa BUZ_ZEMIN ile boyanir; okunacak her sey
     (kart, tablo govdesi, giris alani) beyaz bir yuzeye oturur. Katman farki
     kutu cizmeden hiyerarsi kurar.
  2. CIZGI ince ve soluk. Kalin kenarlik yerine hairline; ayirma isini
     bosluk yapar, cizgi yalnizca hatirlatir.
  3. BOSLUK ucuzdur. Satir yuksekligi ve dar ara sutunlar arayuzun nefes
     almasini saglar; kalabalik bir ekrani hicbir renk kurtarmaz.
"""

from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

from tasarim import (DURUM_RENK, FONT_AILE, FONT_BASLIK_AILE, OLCU,
                     PT, RENK, URUN_ADI)


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
    ilk = f"{get_column_letter(c1)}{r1}"
    if r2 is None or c2 is None:
        return ilk
    return f"{ilk}:{get_column_letter(c2)}{r2}"


def blok_doldur(ws, r1, c1, r2, c2, hex_kod):
    f = dolgu(hex_kod)
    for h in aralik(ws, r1, c1, r2, c2):
        h.fill = f


def birlestir(ws, r1, c1, r2, c2):
    """Hucreleri birlestirir ve sol ust hucreyi dondurur.

    1x1 aralikta birlestirme YAPILMAZ: tek hucrelik bir merge, VBA tarafinda
    ClearContents cagrisini "bu islemi birlestirilmis bir hucrede yapamayiz"
    hatasina dusuruyor.
    """
    if (r1, c1) != (r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    return ws.cell(row=r1, column=c1)


def cerceve(ws, r1, c1, r2, c2, hex_kod=RENK["CIZGI_GRI"], golge_hex=None, stil="thin"):
    """Bir blogun cevresine hairline cerceve cizer.

    golge_hex verilirse ALT kenar bir ton koyu olur. Excel'de hucrelerin
    golgesi yoktur; "isik ustten geliyor" hissi bu tek tonluk farkla verilir.
    Eskiden alt/sag kenar "medium" kalinliktaydi; kalin kenarlik yuzey degil
    cerceve gibi durdugu icin hairline'a indirildi.
    """
    ince = kenar(hex_kod, stil)
    koyu = kenar(golge_hex, stil) if golge_hex else ince
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = Border(
                top=ince if r == r1 else None,
                left=ince if c == c1 else None,
                bottom=koyu if r == r2 else None,
                right=ince if c == c2 else None,
            )


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
    """Her sayfanin ustundeki lacivert bant (satir 1-2) ve vurgu cizgisi (3).

    Solda sistemin adi, sagda bulundugunuz ekran. Kullanicinin nerede oldugunu
    bir bakista anlamasini saglar. Bandin ust satiri alt satirdan kisadir:
    metin optik olarak biraz asagida durur, uygulama baslik cubuklarindaki
    dengeye yaklasir.
    """
    ws.row_dimensions[1].height = OLCU["MASTHEAD_UST"]
    ws.row_dimensions[2].height = OLCU["MASTHEAD_ALT"]
    blok_doldur(ws, 1, 1, 2, son_sutun, RENK["ANA_LACIVERT"])

    # Bant ikiye bolunur ama YARI YARIYA degil: sagdaki yalnizca kisa bir sayfa
    # adi tasir (9-10 karakter), soldaki ise urun adinin tamamini. Yariya
    # bolunseydi uzun bir urun adi kirpilirdi.
    sag_bas = max(3, son_sutun - 2)
    sol = birlestir(ws, 1, 2, 2, sag_bas - 1)
    sol.value = ust_metin
    sol.font = yazi(RENK["BEYAZ"], PT["MASTHEAD_BASLIK"], kalin=True, aile=FONT_BASLIK_AILE)
    sol.alignment = hiza("left", "center", girinti=1)

    sag = birlestir(ws, 1, sag_bas, 2, son_sutun - 1)
    sag.value = buyuk(sayfa_adi)
    sag.font = yazi(RENK["SOLUK_MAVI"], PT["MASTHEAD_ALT"], kalin=True)
    sag.alignment = hiza("right", "center")

    # Bandin altina ince bir vurgu cizgisi
    ws.row_dimensions[3].height = OLCU["MASTHEAD_VURGU"]
    blok_doldur(ws, 3, 1, 3, son_sutun, RENK["KURUMSAL_MAVI"])


def sayfa_basligi(ws, satir, c1, c2, metin, alt_metin=""):
    """Ekranin icindeki buyuk baslik (giris ekranlari icin)."""
    ws.row_dimensions[satir].height = 32.0
    h = birlestir(ws, satir, c1, satir, c2)
    h.value = metin
    h.font = yazi(RENK["ANA_LACIVERT"], PT["EKRAN_BASLIK"], kalin=True, aile=FONT_BASLIK_AILE)
    h.alignment = hiza("left", "center", girinti=1)
    if alt_metin:
        ws.row_dimensions[satir + 1].height = 20.0
        a = birlestir(ws, satir + 1, c1, satir + 1, c2)
        a.value = alt_metin
        a.font = yazi(RENK["METIN_GRI"], PT["EKRAN_ALT"])
        a.alignment = hiza("left", "center", girinti=1)
    return h


def bolum_basligi(ws, satir, c1, c2, metin, alt_metin=""):
    """Numarali bolum basligi."""
    ws.row_dimensions[satir].height = 22.0
    h = birlestir(ws, satir, c1, satir, c2)
    h.value = metin
    h.font = yazi(RENK["ANA_LACIVERT"], PT["BOLUM_BASLIK"], kalin=True, aile=FONT_BASLIK_AILE)
    h.alignment = hiza("left", "bottom", girinti=1)
    if alt_metin:
        ws.row_dimensions[satir + 1].height = 15.0
        a = birlestir(ws, satir + 1, c1, satir + 1, c2)
        a.value = alt_metin
        a.font = yazi(RENK["METIN_SOLUK"], PT["BOLUM_ALT"])
        a.alignment = hiza("left", "top", girinti=1)


def etiket(ws, satir, sutun, metin, zorunlu=False):
    """Alan ustundeki mikro etiket. Buyuk harfe cevirme burada yapilir."""
    h = ws.cell(row=satir, column=sutun)
    h.value = buyuk(metin) + (" *" if zorunlu else "")
    h.font = yazi(RENK["METIN_GRI"], PT["ETIKET"], kalin=True)
    h.alignment = hiza("left", "center")
    return h


def form_alani(ws, wb, satir, c1, c2, ad, sayfa_adi, yukseklik=None,
               coklu_satir=False):
    """Beyaz zeminli, alt kenarligi olan, KILITSIZ bir giris alani.

    Sayfa korumali oldugu icin yalnizca bu hucreler yazilabilir; kullanici
    tasarimi yanlislikla bozamaz. Alan adlandirilmis aralik olarak kaydedilir,
    VBA hucre adresi degil bu adi kullanir.
    """
    ws.row_dimensions[satir].height = yukseklik or OLCU["SATIR_ALAN"]

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

    wb.defined_names.add(
        DefinedName(ad, attr_text=f"'{sayfa_adi}'!${get_column_letter(c1)}${satir}")
    )
    return h


def ipucu_satiri(ws, satir, c1, c2, metin):
    ws.row_dimensions[satir].height = 14.0
    h = birlestir(ws, satir, c1, satir, c2)
    h.value = metin
    h.font = yazi(RENK["METIN_SOLUK"], PT["IPUCU"], italik=True)
    h.alignment = hiza("left", "center", girinti=1)
    return h


def not_satiri(ws, satir, c1, c2, metin):
    """Ekranin altindaki kucuk aciklama satiri."""
    h = birlestir(ws, satir, c1, satir, c2)
    h.value = metin
    h.font = yazi(RENK["METIN_SOLUK"], PT["NOT"], italik=True)
    h.alignment = hiza("left", "center", girinti=1)
    return h


def bosluk(ws, satir, yukseklik=OLCU["SATIR_BOSLUK"]):
    ws.row_dimensions[satir].height = yukseklik


def ayrac(ws, satir, c1, c2, hex_kod=RENK["CIZGI_GRI"]):
    """1 punto yuksekliginde dolu bir satir -- yatay ayrac cizgisi.

    Excel'de iki blok arasina cizgi cekmenin en temiz yolu kenarlik degil,
    incecik bir satirin kendisini boyamaktir: kenarlik yakinlastirinca
    kalinlasir, bu yontem her yakinlastirmada ayni kalir.
    """
    ws.row_dimensions[satir].height = 1.0
    blok_doldur(ws, satir, c1, satir, c2, hex_kod)


# --------------------------------------------------------------------------
# Kart
# --------------------------------------------------------------------------
def kart(ws, r1, c1, r2, c2, vurgu_hex=None, zemin_hex=RENK["BEYAZ"]):
    """Beyaz yuzey + hairline cerceve. Icerik bunun uzerine yazilir.

    vurgu_hex verilirse sol kenar renkli ve kalin olur ("accent card"):
    bir bloga anlam yuklemenin baslik yazmadan en ucuz yolu.
    """
    blok_doldur(ws, r1, c1, r2, c2, zemin_hex)
    cerceve(ws, r1, c1, r2, c2, RENK["CIZGI_GRI"], golge_hex=RENK["GOLGE_GRI"])
    if vurgu_hex:
        sol = kenar(vurgu_hex, "thick")
        for r in range(r1, r2 + 1):
            h = ws.cell(row=r, column=c1)
            m = h.border
            h.border = Border(top=m.top, bottom=m.bottom, right=m.right, left=sol)


def kpi_karti(ws, wb, r1, c1, c2, etiket_metni, ad, sayfa_adi,
              alt_metin="", sayi_bicimi="#,##0", vurgu_hex=None, yuzey=True):
    """Uc satirlik gosterge: ustte kucuk etiket, ortada buyuk rakam, altta not.

    r1 kartin ilk satiridir; kart r1..r1+2 arasini kaplar. Deger hucresi
    adlandirilmis aralik olur, modPano oraya yazar.

    yuzey=False: beyaz zemin ve cerceve cizilmez. Pano'da kartin govdesi COM
    asamasinda eklenen yuvarlak kose bir SEKILDIR; sekil hucrenin ustunu
    kapattigi icin altina ikinci bir yuzey cizmek gereksizdir. Sekil bir
    sebeple olusmazsa gosterge duz bir zemin uzerinde yine de okunur kalir.
    """
    vurgu = vurgu_hex or RENK["ANA_LACIVERT"]

    ws.row_dimensions[r1].height = 16.0
    ws.row_dimensions[r1 + 1].height = 42.0
    ws.row_dimensions[r1 + 2].height = 22.0

    if yuzey:
        kart(ws, r1, c1, r1 + 2, c2, vurgu_hex=vurgu)

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
    a.font = yazi(RENK["METIN_SOLUK"], PT["KPI_ALT"])
    a.alignment = hiza("left", "top", girinti=1)

    wb.defined_names.add(
        DefinedName(ad, attr_text=f"'{sayfa_adi}'!${get_column_letter(c1)}${r1 + 1}")
    )
    return d


# --------------------------------------------------------------------------
# Tablo
# --------------------------------------------------------------------------
def tablo_basligi(ws, satir, c1, c2, basliklar, ortali=()):
    """Acik zeminli veri tablosu basligi.

    Eskiden baslik satiri dolu laciverttti. Koyu bant, tablonun kendisinden
    daha cok dikkat cekiyordu; modern veri tablolarindaki gibi baslik artik
    beyaz zemin uzerinde kucuk, buyuk harf ve gri. Tablonun nerede basladigini
    satirin altindaki lacivert cizgi soyluyor.
    """
    ws.row_dimensions[satir].height = 28.0
    blok_doldur(ws, satir, c1, satir, c2, RENK["BEYAZ"])
    alt = kenar(RENK["ANA_LACIVERT"])
    for i, metin in enumerate(basliklar):
        c = c1 + i
        h = ws.cell(row=satir, column=c)
        h.value = buyuk(metin)
        h.font = yazi(RENK["METIN_GRI"], PT["TABLO_BASLIK"], kalin=True)
        h.alignment = hiza("center" if c in ortali else "left", "center",
                           kaydir=True, girinti=0 if c in ortali else 1)
    for c in range(c1, c2 + 1):
        ws.cell(row=satir, column=c).border = Border(bottom=alt)


def tablo_govde_stili(ws, r1, r2, c1, c2, serit=False, ortali=()):
    """Tablo govdesi: beyaz yuzey, satir altinda hairline, dikey cizgi yok.

    Zebra serit varsayilan olarak kapalidir. Iki renkli satirlar, satir sayisi
    az oldugunda gorsel gurultu yaratiyordu; ayirma isini satir altindaki
    incecik cizgi zaten yapiyor.
    """
    ince = kenar(RENK["CIZGI_INCE"])
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].height = OLCU["SATIR_TABLO"]
        zemin = RENK["BEYAZ"]
        if serit and (r - r1) % 2 == 1:
            zemin = RENK["BUZ_ZEMIN"]
        for c in range(c1, c2 + 1):
            h = ws.cell(row=r, column=c)
            h.fill = dolgu(zemin)
            h.border = Border(bottom=ince)
            h.font = yazi(RENK["METIN_KOYU"], PT["GOVDE"])
            h.alignment = hiza("center" if c in ortali else "left", "center",
                               girinti=0 if c in ortali else 1)


def durum_kosullu_bicim(ws, hucre_araligi):
    """Durum sutununa sekiz durumun rozet renklerini kosullu bicim olarak kurar.

    Kosullu bicim, VBA tabloyu her yeniden kurdugunda renkleri tekrar boyamak
    zorunda kalmasini onler.
    """
    from openpyxl.formatting.rule import CellIsRule

    for durum, (zemin, yazi_renk) in DURUM_RENK.items():
        ws.conditional_formatting.add(
            hucre_araligi,
            CellIsRule(operator="equal", formula=[f'"{durum}"'],
                       fill=dolgu(zemin),
                       font=yazi(yazi_renk, PT["GOVDE"], kalin=True)),
        )


# --------------------------------------------------------------------------
# Bilgi kutusu
# --------------------------------------------------------------------------
def bilgi_kutusu(ws, r1, c1, r2, c2, metin, zemin=RENK["ACIK_MAVI"],
                 yazi_hex=RENK["ANA_LACIVERT"], vurgu=RENK["KURUMSAL_MAVI"]):
    """Sol kenarinda renkli serit olan aciklama blogu (web'deki "callout").

    Cerceve yerine tek bir dikey serit kullanilir: goz metne gider, kutuya
    degil.
    """
    blok_doldur(ws, r1, c1, r2, c2, zemin)
    serit = kenar(vurgu, "thick")
    for r in range(r1, r2 + 1):
        ws.cell(row=r, column=c1).border = Border(left=serit)

    h = birlestir(ws, r1, c1, r2, c2)
    h.value = metin
    h.font = yazi(yazi_hex, PT["GOVDE"])
    h.alignment = hiza("left", "center", kaydir=True, girinti=1)
    return h


# --------------------------------------------------------------------------
# Listeler sayfasi
# --------------------------------------------------------------------------
def listeler_sayfasi_kur(wb, listeler):
    """Veri dogrulama listelerini tasiyan gizli sayfa."""
    ws = wb.create_sheet("Listeler")
    for i, (ad, degerler) in enumerate(listeler.items(), start=1):
        harf = get_column_letter(i)
        b = ws.cell(row=1, column=i, value=ad)
        b.font = yazi(RENK["METIN_GRI"], PT["NOT"], kalin=True)
        for j, deger in enumerate(degerler, start=2):
            ws.cell(row=j, column=i, value=deger)
        ws.column_dimensions[harf].width = 24
        wb.defined_names.add(DefinedName(
            f"lst_{ad}",
            attr_text=f"Listeler!${harf}$2:${harf}${len(degerler) + 1}"))
    ws.sheet_state = "hidden"
    return ws
