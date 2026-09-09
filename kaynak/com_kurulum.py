# -*- coding: utf-8 -*-
"""COM asamasi: VBA enjeksiyonu, dugmeler, grafikler, koruma, .xlsm kaydi.

openpyxl gorsel iskeleti kurar ama uc seyi yapamaz: VBA projesi yazamaz,
yuvarlatilmis dugme sekli ekleyip makroya baglayamaz, canli grafik kuramaz.
Bu modul o uc isi gercek Excel uzerinden tamamlar.

KODLAMA NOTU -- onemli:
VBComponents.Import bir dosyayi sistemin ANSI kod sayfasiyla okur; Turkce
karakterler baska kod sayfasina sahip bir makinede bozulurdu. Bu yuzden
Import HIC kullanilmaz. Moduller bos olarak olusturulup kaynak metin
CodeModule.AddFromString ile verilir; bu yol Unicode BSTR uzerinden gecer
ve kod sayfasindan tamamen bagimsizdir.
"""

import os
import re
import winreg
from contextlib import contextmanager

import pythoncom
import win32com.client

from tasarim import (DUGME, FONT_AILE, FONT_BASLIK_AILE, KART, PT, RENK,
                     rgb_long)

# --- Excel sabitleri (surumden bagimsiz olsun diye sayisal) ---------------
XL_OPENXML_MACRO = 52          # xlOpenXMLWorkbookMacroEnabled (.xlsm)
XL_OPENXML = 51                # xlOpenXMLWorkbook (.xlsx)
VBEXT_CT_STD_MODULE = 1
MSO_SHAPE_ROUNDED_RECT = 5
MSO_TEXT_ORIENT_HORIZ = 1
MSO_TRUE, MSO_FALSE = -1, 0
MSO_ANCHOR_TOP = 1
MSO_ANCHOR_MIDDLE = 3
MSO_ALIGN_LEFT = 1
MSO_ALIGN_CENTER = 2
MSO_SHADOW_DIS_ALT_ORTA = 25   # msoShadow25 -- alt-orta dis golge
XL_FREE_FLOATING = 3           # xlFreeFloating
XL_YUKARI = -4162              # xlUp
XL_SOLA = -4159                # xlToLeft
MSO_GUVENLIK_KAPALI = 3        # msoAutomationSecurityForceDisable


# ==========================================================================
#  AccessVBOM -- "VBA proje nesne modeline erisime guven" ayari
# ==========================================================================
@contextmanager
def vba_erisimi_acik(surumler=("16.0", "15.0", "14.0")):
    """Ayari gecici olarak acar ve NE OLURSA OLSUN eski degerine dondurur.

    Ayar Excel baslatilmadan once yazilmalidir; calisan bir Excel bu degeri
    yeniden okumaz. Bu yuzden context manager Excel'i saran en dis katmandir.
    """
    onceki = []
    try:
        for surum in surumler:
            yol = rf"Software\Microsoft\Office\{surum}\Excel\Security"
            try:
                anahtar = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, yol, 0,
                                             winreg.KEY_READ | winreg.KEY_WRITE)
            except OSError:
                continue
            try:
                eski = winreg.QueryValueEx(anahtar, "AccessVBOM")[0]
            except OSError:
                eski = None
            onceki.append((anahtar, eski))
            winreg.SetValueEx(anahtar, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        yield
    finally:
        for anahtar, eski in onceki:
            try:
                if eski is None:
                    winreg.DeleteValue(anahtar, "AccessVBOM")
                else:
                    winreg.SetValueEx(anahtar, "AccessVBOM", 0, winreg.REG_DWORD, eski)
            except OSError:
                pass
            finally:
                anahtar.Close()


@contextmanager
def excel_oturumu(gorunur=False):
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = gorunur
    app.DisplayAlerts = False
    app.EnableEvents = False          # uretim sirasinda Workbook_Open calismasin
    app.ScreenUpdating = gorunur
    # EnableEvents tek basina yetmez: uretilen kitap yeniden acildiginda
    # (kitap_dogrula, veri_aktar) makrolarin hic calismamasi gerekir. Yonetim
    # kitabinin Workbook_Open'i yedek alip kitabi salt okunura cevirir; uretim
    # sirasinda ikisi de istenmez.
    try:
        app.AutomationSecurity = MSO_GUVENLIK_KAPALI
    except Exception:
        pass
    try:
        yield app
    finally:
        try:
            app.DisplayAlerts = True
            app.Quit()
        except Exception:
            pass
        del app
        pythoncom.CoUninitialize()


# ==========================================================================
#  VBA enjeksiyonu
# ==========================================================================
_AD_DESENI = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"\s*$',
                        re.IGNORECASE | re.MULTILINE)


def bas_ayristir(bas_yolu):
    """.bas dosyasini (modul adi, kaynak metin) olarak dondurur.

    Attribute satiri yalnizca Import icin anlamlidir; AddFromString yolunda
    kaynaga girmemesi gerekir, modul adini oradan alip satiri atariz.
    """
    with open(bas_yolu, "r", encoding="utf-8") as f:
        ham = f.read()

    eslesme = _AD_DESENI.search(ham)
    if not eslesme:
        raise ValueError(f'{bas_yolu}: Attribute VB_Name satırı yok.')
    ad = eslesme.group(1)
    kaynak = _AD_DESENI.sub("", ham, count=1).lstrip("\r\n")
    return ad, kaynak


def modul_ekle(wb, bas_yolu):
    ad, kaynak = bas_ayristir(bas_yolu)
    bilesen = wb.VBProject.VBComponents.Add(VBEXT_CT_STD_MODULE)
    bilesen.Name = ad
    bilesen.CodeModule.AddFromString(kaynak)
    return ad


def thisworkbook_kodu_yaz(wb, kaynak):
    # Bilesen adi "ThisWorkbook" degil, kitabin CodeName'idir; Excel'in dil
    # surumune gore yerellesir (Turkce'de "BuÇalışmaKitabı"). Sabit ad yerine
    # kitabin kendi CodeName'i kullanilir.
    modul = wb.VBProject.VBComponents(wb.CodeName).CodeModule
    if modul.CountOfLines > 0:
        modul.DeleteLines(1, modul.CountOfLines)
    modul.AddFromString(kaynak)


# ==========================================================================
#  Sekil yardimcilari
# ==========================================================================
def hucre_kutusu(ws, r1, c1, r2, c2, yatay_inset=0.0, dikey_inset=0.0):
    """Bir hucre blogunun nokta cinsinden dikdortgenini dondurur.

    Sekiller hucrelere degil sayfaya, nokta cinsinden yerlestirilir. Bu
    fonksiyon iki dunyayi birbirine baglar: duzen hucre izgarasinda tasarlanir,
    sekil o izgaranin olcusunu alir.
    """
    ust_sol = ws.Cells(r1, c1)
    alt_sag = ws.Cells(r2, c2)
    sol = ust_sol.Left + yatay_inset
    ust = ust_sol.Top + dikey_inset
    sag = alt_sag.Left + alt_sag.Width - yatay_inset
    alt = alt_sag.Top + alt_sag.Height - dikey_inset
    return sol, ust, sag - sol, alt - ust


def _golge(sekil, bulanik, kaydirma, seffaflik, hex_kod):
    """Yumusak dis golge. Golge kozmetiktir; basarisiz olursa is durmaz."""
    try:
        sekil.Shadow.Type = MSO_SHADOW_DIS_ALT_ORTA
        sekil.Shadow.Visible = MSO_TRUE
        sekil.Shadow.Blur = bulanik
        sekil.Shadow.OffsetX = 0
        sekil.Shadow.OffsetY = kaydirma
        sekil.Shadow.Transparency = seffaflik
        sekil.Shadow.ForeColor.RGB = rgb_long(hex_kod)
    except Exception:
        pass


def kart_govdesi(ws, sol, ust, gen, yuk, ad):
    """Yuvarlak koseli, beyaz, yumusak golgeli kart/panel govdesi.

    Excel hucrelere yuvarlak kose ve golge veremez; sekiller verebilir. Bedeli
    su: sekiller HER ZAMAN hucrelerin ustunde cizilir, yani kartin metni de
    sekle tasinmak zorundadir (bkz. metin_katmani).
    """
    s = ws.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECT, sol, ust, gen, yuk)
    s.Name = ad
    s.Placement = XL_FREE_FLOATING
    try:
        s.Adjustments[1] = KART["KOSE_ORAN"]
    except Exception:
        pass
    s.Fill.Visible = MSO_TRUE
    s.Fill.ForeColor.RGB = rgb_long(RENK["BEYAZ"])
    s.Line.Visible = MSO_TRUE
    s.Line.ForeColor.RGB = rgb_long(RENK["CIZGI_GRI"])
    s.Line.Weight = KART["CIZGI_KALINLIK"]
    _golge(s, KART["GOLGE_BULANIK"], KART["GOLGE_KAYDIRMA"],
           KART["GOLGE_SEFFAF"], RENK["DERIN_LACIVERT"])
    return s


def metin_katmani(ws, sol, ust, gen, yuk, ad, metin, punto,
                  renk=RENK["METIN_KOYU"], aile=FONT_AILE):
    """Kartin uzerine oturan dolgusuz/cizgisiz metin kutusu.

    Ayri bir sekil olmasinin sebebi: tek bir sekilde farkli punto ve renkte
    uc satir tutmak, VBA metni tazeledikce bicimin dagilmasi demektir. Uc ayri
    katman, her birinin bicimini kendi uzerinde tutar.
    """
    tb = ws.Shapes.AddTextbox(MSO_TEXT_ORIENT_HORIZ, sol, ust, gen, yuk)
    tb.Name = ad
    tb.Placement = XL_FREE_FLOATING
    tb.Fill.Visible = MSO_FALSE
    tb.Line.Visible = MSO_FALSE

    tf = tb.TextFrame2
    tf.VerticalAnchor = MSO_ANCHOR_TOP
    tf.WordWrap = MSO_TRUE
    tf.AutoSize = 0
    tf.MarginLeft = 0
    tf.MarginRight = 0
    tf.MarginTop = 0
    tf.MarginBottom = 0

    tr = tf.TextRange
    tr.Text = metin
    tr.ParagraphFormat.Alignment = MSO_ALIGN_LEFT
    tr.Font.Name = aile
    tr.Font.Size = punto
    tr.Font.Bold = MSO_FALSE
    tr.Font.Fill.ForeColor.RGB = rgb_long(renk)
    return tb


def kpi_karti_ciz(ws, r1, c1, c2, etiket, alt_metin, vurgu_hex, ad):
    """Uc metin katmanli, gercek yuvarlak koseli KPI karti.

    Deger katmaninin adi "kpi_deger_<ad>"dir; modPano her yenilemede oraya
    yazar. Sayinin kendisi ayrica altindaki hucrede durur -- gostergelerin tek
    dogruluk kaynagi o hucredir, sekil yalnizca gorunumdur.
    """
    from uret_ortak import buyuk

    sol, ust, gen, yuk = hucre_kutusu(ws, r1, c1, r1 + 2, c2,
                                      yatay_inset=KART["YATAY_INSET"])
    kart_govdesi(ws, sol, ust, gen, yuk, f"kart_{ad}")

    # Uc katman ust uste BINMEMELIDIR: metin kutulari saydamdir, cakisirlarsa
    # iki satir ayni yerde okunur. Dikey ritim burada elle kurulur ve kartin
    # hucre yuksekligiyle (16 + 42 + 22 = 80pt) birlikte degisir.
    p = KART["IC_BOSLUK"]
    ic_gen = gen - 2 * p
    metin_katmani(ws, sol + p, ust + 11, ic_gen, 13, f"kpi_etiket_{ad}",
                  buyuk(etiket), PT["KPI_ETIKET"], RENK["METIN_GRI"],
                  FONT_BASLIK_AILE)
    metin_katmani(ws, sol + p, ust + 26, ic_gen, 38, f"kpi_deger_{ad}",
                  "0", PT["KPI_RAKAM"], vurgu_hex, FONT_BASLIK_AILE)
    metin_katmani(ws, sol + p, ust + 64, ic_gen, 13, f"kpi_alt_{ad}",
                  alt_metin, PT["KPI_ALT"], RENK["METIN_SOLUK"])


# ==========================================================================
#  Dugmeler
# ==========================================================================
#  Uc varyant: birincil (dolu lacivert), ikincil (beyaz + ince kenarlik),
#  sessiz (dolgusuz, yalnizca yazi). Sessiz varyant "Cikis" gibi ikincil
#  eylemler icindir: ekranda yer kaplar ama dikkat calmaz.
def dugme_ekle(ws, hucre_adresi, metin, makro, birincil=True, varyant=None,
               genislik=None, yukseklik=None, sol_kaydir=0.0):
    """Yuvarlatilmis dikdortgen dugme ekler ve makroya baglar.

    Excel'in kendi Forms dugmesi kurumsal bir arayuz icin fazla eski gorunur;
    sekil tabanli dugme renk, kose yaricapi ve tipografi kontrolu verir.
    """
    if varyant is None:
        varyant = "birincil" if birincil else "ikincil"

    hucre = ws.Range(hucre_adresi)
    g = genislik or DUGME["GENISLIK"]
    y = yukseklik or DUGME["YUKSEKLIK"]
    sol = hucre.Left + sol_kaydir
    ust = hucre.Top + max(0.0, (hucre.Height - y) / 2.0)

    sekil = ws.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECT, sol, ust, g, y)
    sekil.Name = "btn_" + makro.split(".")[-1]
    sekil.Placement = XL_FREE_FLOATING

    try:
        sekil.Adjustments[1] = DUGME["KOSE_ORAN"]
    except Exception:
        pass          # kose yaricapi kozmetiktir; varsayilan da kabul edilebilir

    if varyant == "birincil":
        sekil.Fill.Visible = MSO_TRUE
        sekil.Fill.ForeColor.RGB = rgb_long(RENK["ANA_LACIVERT"])
        sekil.Line.Visible = MSO_FALSE
        yazi_rengi = RENK["BEYAZ"]
        golgeli = True
    elif varyant == "sessiz":
        sekil.Fill.Visible = MSO_FALSE
        sekil.Line.Visible = MSO_FALSE
        yazi_rengi = RENK["METIN_GRI"]
        golgeli = False
    else:                                   # ikincil
        sekil.Fill.Visible = MSO_TRUE
        sekil.Fill.ForeColor.RGB = rgb_long(RENK["BEYAZ"])
        sekil.Line.Visible = MSO_TRUE
        sekil.Line.ForeColor.RGB = rgb_long(RENK["GOLGE_GRI"])
        sekil.Line.Weight = 0.75
        yazi_rengi = RENK["ANA_LACIVERT"]
        golgeli = True

    tf = sekil.TextFrame2
    tf.VerticalAnchor = MSO_ANCHOR_MIDDLE
    tf.MarginLeft = 8
    tf.MarginRight = 8
    tf.MarginTop = 0
    tf.MarginBottom = 0
    tr = tf.TextRange
    tr.Text = metin
    tr.ParagraphFormat.Alignment = MSO_ALIGN_CENTER
    tr.Font.Name = FONT_BASLIK_AILE
    tr.Font.Size = PT["DUGME"]
    # Semibold zaten agir bir kesim; ustune Bold vermek Excel'e sahte kalinlik
    # urettiriyor ve harfler bulaniyor.
    tr.Font.Bold = MSO_FALSE
    tr.Font.Fill.ForeColor.RGB = rgb_long(yazi_rengi)

    if golgeli:
        _golge(sekil, DUGME["GOLGE_BULANIK"], DUGME["GOLGE_KAYDIRMA"],
               DUGME["GOLGE_SEFFAF"], RENK["DERIN_LACIVERT"])

    sekil.OnAction = makro
    return sekil


# ==========================================================================
#  Ana islem
# ==========================================================================
def veri_aktar(app, wb, eski_yol, sifre, sayfalar):
    """Eski kitaptaki depo satirlarini yeni uretilen kitaba tasir.

    Yonetim kitabi artik VERI DEPOSUDUR: yeniden uretmek, hicbir sey
    yapilmazsa butun onerileri ve butun degerlendirme gecmisini silmek
    demektir. Bu yordam onlari once okuyup yenisine yazar; boylece ekran
    degisikligi icin kur.py'yi yeniden calistirmak guvenli kalir.

    Kaydedilen sayfa sayilarini {sayfa: satir} olarak dondurur.
    """
    eski = app.Workbooks.Open(os.path.abspath(eski_yol), 0, True, None, sifre or "",
                              "", True)
    try:
        toplam = {}
        for ad in sayfalar:
            try:
                kaynak = eski.Worksheets(ad)
            except Exception:
                toplam[ad] = 0        # eski surumde bu sayfa yoktu
                continue

            son = kaynak.Cells(kaynak.Rows.Count, 1).End(XL_YUKARI).Row
            if son < 2:
                toplam[ad] = 0
                continue
            sutun = kaynak.Cells(1, kaynak.Columns.Count).End(XL_SOLA).Column

            deger = kaynak.Range(kaynak.Cells(2, 1),
                                 kaynak.Cells(son, sutun)).Value
            hedef = wb.Worksheets(ad)
            alan = hedef.Range(hedef.Cells(2, 1), hedef.Cells(son, sutun))
            alan.NumberFormat = "@"
            alan.Value = deger
            toplam[ad] = son - 1
        return toplam
    finally:
        try:
            eski.Close(SaveChanges=False)
        except Exception:
            pass


def kitap_isle(app, taslak_yol, hedef_yol, modul_yollari, thisworkbook_kod,
               dugmeler, ek_islem=None, koruma_sifresi=None,
               dosya_sifresi=None, veri_kaynagi=None, veri_sayfalari=()):
    """Taslak .xlsx dosyasini alir, VBA + dugme + koruma ekleyip .xlsm kaydeder.

    dosya_sifresi verilirse dosya ACILIS PAROLASIYLA sifrelenir; yonetim
    kitabinin gizliligi buna dayanir.

    veri_kaynagi verilirse oradaki depo satirlari yeni kitaba tasinir
    (bkz. veri_aktar).
    """
    wb = app.Workbooks.Open(os.path.abspath(taslak_yol))
    try:
        eklenen = [modul_ekle(wb, y) for y in modul_yollari]
        thisworkbook_kodu_yaz(wb, thisworkbook_kod)

        for d in dugmeler:
            ws = wb.Worksheets(d["sayfa"])
            gorunurluk = ws.Visible
            ws.Visible = -1                      # sekil eklemek icin gorunur olmali
            dugme_ekle(ws, d["hucre"], d["metin"], d["makro"],
                       birincil=d.get("birincil", True),
                       varyant=d.get("varyant"),
                       genislik=d.get("genislik"),
                       yukseklik=d.get("yukseklik"),
                       sol_kaydir=d.get("sol_kaydir", 0.0))
            ws.Visible = gorunurluk

        if ek_islem:
            ek_islem(wb)

        # Veri aktarimi korumadan ONCE yapilir: sonra yapilsa her sayfayi
        # yeniden acmak gerekirdi.
        if veri_kaynagi and veri_sayfalari:
            aktarilan = veri_aktar(app, wb, veri_kaynagi, dosya_sifresi,
                                   veri_sayfalari)
            print("        veri taşındı: " +
                  ", ".join(f"{a}={n}" for a, n in aktarilan.items()))

        if koruma_sifresi:
            for ws in wb.Worksheets:
                gorunurluk = ws.Visible
                ws.Visible = -1
                ws.Protect(Password=koruma_sifresi, DrawingObjects=False,
                           Contents=True, Scenarios=False,
                           AllowFiltering=True, AllowSorting=True)
                ws.Visible = gorunurluk

        # Ilk gorunur sayfayi etkin birak ki kitap dogru ekranla acilsin.
        for ws in wb.Worksheets:
            if ws.Visible == -1:
                ws.Activate()
                break

        hedef = os.path.abspath(hedef_yol)
        klasor = os.path.dirname(hedef)
        if klasor and not os.path.isdir(klasor):
            os.makedirs(klasor, exist_ok=True)
        if os.path.exists(hedef):
            os.remove(hedef)
        if dosya_sifresi:
            wb.SaveAs(hedef, FileFormat=XL_OPENXML_MACRO, Password=dosya_sifresi)
        else:
            wb.SaveAs(hedef, FileFormat=XL_OPENXML_MACRO)
        return eklenen
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass


def kitap_dogrula(app, xlsm_yolu, beklenen_moduller, beklenen_makrolar,
                  dosya_sifresi=None):
    """Uretilen dosyayi yeniden acip modul ve dugme baglantilarini denetler.

    Parola HER ZAMAN acikca gecilir: parolasi verilmeyen sifreli bir dosya
    gorunmez Excel'de parola diyalogu acar ve cagri hic geri donmez.
    """
    sorunlar = []
    wb = app.Workbooks.Open(os.path.abspath(xlsm_yolu), 0, False, None,
                            dosya_sifresi or "", "", True)
    try:
        mevcut = {c.Name for c in wb.VBProject.VBComponents}
        for m in beklenen_moduller:
            if m not in mevcut:
                sorunlar.append(f"Modül eksik: {m}")

        bagli = set()
        for ws in wb.Worksheets:
            for sekil in ws.Shapes:
                try:
                    if sekil.OnAction:
                        bagli.add(sekil.OnAction)
                except Exception:
                    pass
        for makro in beklenen_makrolar:
            if makro not in bagli:
                sorunlar.append(f"Düğme bağlı değil: {makro}")

        return sorunlar
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass
