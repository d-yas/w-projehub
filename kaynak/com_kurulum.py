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

from tasarim import DUGME, RENK, rgb_long

# --- Excel sabitleri (surumden bagimsiz olsun diye sayisal) ---------------
XL_OPENXML_MACRO = 52          # xlOpenXMLWorkbookMacroEnabled (.xlsm)
XL_OPENXML = 51                # xlOpenXMLWorkbook (.xlsx)
VBEXT_CT_STD_MODULE = 1
MSO_SHAPE_ROUNDED_RECT = 5
MSO_TRUE, MSO_FALSE = -1, 0
MSO_ANCHOR_MIDDLE = 3
MSO_ALIGN_CENTER = 2
XL_FREE_FLOATING = 3           # xlFreeFloating


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
#  Dugmeler
# ==========================================================================
def dugme_ekle(ws, hucre_adresi, metin, makro, birincil=True,
               genislik=None, yukseklik=None, sol_kaydir=0.0):
    """Yuvarlatilmis dikdortgen dugme ekler ve makroya baglar.

    Excel'in kendi Forms dugmesi kurumsal bir arayuz icin fazla eski gorunur;
    sekil tabanli dugme renk, kose yaricapi ve tipografi kontrolu verir.
    """
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

    if birincil:
        sekil.Fill.Visible = MSO_TRUE
        sekil.Fill.ForeColor.RGB = rgb_long(RENK["ANA_LACIVERT"])
        sekil.Line.Visible = MSO_FALSE
        yazi_rengi = rgb_long(RENK["BEYAZ"])
    else:
        sekil.Fill.Visible = MSO_TRUE
        sekil.Fill.ForeColor.RGB = rgb_long(RENK["BEYAZ"])
        sekil.Line.Visible = MSO_TRUE
        sekil.Line.ForeColor.RGB = rgb_long(RENK["ANA_LACIVERT"])
        sekil.Line.Weight = 1.0
        yazi_rengi = rgb_long(RENK["ANA_LACIVERT"])

    tf = sekil.TextFrame2
    tf.VerticalAnchor = MSO_ANCHOR_MIDDLE
    tf.MarginLeft = 2
    tf.MarginRight = 2
    tf.MarginTop = 0
    tf.MarginBottom = 0
    tr = tf.TextRange
    tr.Text = metin
    tr.ParagraphFormat.Alignment = MSO_ALIGN_CENTER
    tr.Font.Name = "Segoe UI Semibold"
    tr.Font.Size = 11
    tr.Font.Bold = MSO_TRUE
    tr.Font.Fill.ForeColor.RGB = yazi_rengi

    # Hafif golge: dugmeyi zeminden ayirir, abartili durmaz.
    try:
        sekil.Shadow.Visible = MSO_TRUE
        sekil.Shadow.Style = 1
        sekil.Shadow.Blur = 5
        sekil.Shadow.OffsetX = 0
        sekil.Shadow.OffsetY = 1.5
        sekil.Shadow.Transparency = 0.72
        sekil.Shadow.ForeColor.RGB = rgb_long(RENK["ANA_LACIVERT"])
    except Exception:
        pass

    sekil.OnAction = makro
    return sekil


# ==========================================================================
#  Ana islem
# ==========================================================================
def kitap_isle(app, taslak_yol, hedef_yol, modul_yollari, thisworkbook_kod,
               dugmeler, ek_islem=None, koruma_sifresi=None):
    """Taslak .xlsx dosyasini alir, VBA + dugme + koruma ekleyip .xlsm kaydeder."""
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
                       genislik=d.get("genislik"),
                       yukseklik=d.get("yukseklik"),
                       sol_kaydir=d.get("sol_kaydir", 0.0))
            ws.Visible = gorunurluk

        if ek_islem:
            ek_islem(wb)

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
        wb.SaveAs(hedef, FileFormat=XL_OPENXML_MACRO)
        return eklenen
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass


def kitap_dogrula(app, xlsm_yolu, beklenen_moduller, beklenen_makrolar):
    """Uretilen dosyayi yeniden acip modul ve dugme baglantilarini denetler."""
    sorunlar = []
    wb = app.Workbooks.Open(os.path.abspath(xlsm_yolu))
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
