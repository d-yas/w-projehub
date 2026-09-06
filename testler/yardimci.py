# -*- coding: utf-8 -*-
r"""Test altyapisi: sahte ortak klasor + gercek Excel oturumu.

ONEMLI -- testler neden gecici klasorde calisir:
Proje klasoru OneDrive ile eslenmis olabilir. Eslenmis klasorlerde Excel
kitabin konumunu disk yolu yerine "https://..." adresi olarak bildirir ve
Guvenilir Konum ayari beklendigi gibi davranmaz. Testler bu belirsizligi
disarida birakmak icin %TEMP% altinda -- OneDrive DISINDA -- kendi
"ortak klasorunu" kurar. Uretim ortami zaten bir UNC paylasimi olacaktir.
"""

import os
import shutil
import sys
import tempfile
import threading
import time
from contextlib import contextmanager
from datetime import datetime

import pythoncom
import win32com.client

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CIKTI = os.path.join(KOK, "cikti")

MSO_OTOMASYON_DUSUK = 1        # msoAutomationSecurityLow


class Sonuc:
    """Kucuk bir test toplayici; harici bagimlilik istemiyoruz."""

    def __init__(self, baslik):
        self.baslik = baslik
        self.gecen = 0
        self.kalan = []

    def kontrol(self, ad, kosul, ayrinti=""):
        if kosul:
            self.gecen += 1
        else:
            self.kalan.append(f"{ad}  {ayrinti}".strip())
        return bool(kosul)

    def esit(self, ad, beklenen, bulunan):
        return self.kontrol(ad, beklenen == bulunan,
                            f"(beklenen: {beklenen!r}, bulunan: {bulunan!r})")

    def bitir(self):
        toplam = self.gecen + len(self.kalan)
        print()
        print("=" * 66)
        if self.kalan:
            print(f"  {self.baslik}: {self.gecen}/{toplam} geçti — "
                  f"{len(self.kalan)} BAŞARISIZ")
            for k in self.kalan:
                print(f"    ! {k}")
            print("=" * 66)
            return 1
        print(f"  {self.baslik}: {toplam}/{toplam} kontrolün tamamı geçti")
        print("=" * 66)
        return 0


def yil():
    return datetime.now().year


class Ortam:
    """Gecici bir 'projeoneri\\' paylasim klasoru ve icindeki iki kitap."""

    def __init__(self, kok):
        self.kok = kok
        self.paylasim = os.path.join(kok, "projeoneri")
        self.yonetim = os.path.join(self.paylasim, "yonetim")
        # Gonderimlerin TEK ve kalici yeri. Ayri bir "gelen kutusu" yoktur:
        # personel dogrudan buraya yazar ama icini goremez.
        self.oneriler = os.path.join(self.yonetim, "oneriler")
        self.degerlendirme = os.path.join(self.yonetim, "degerlendirme")
        self.oneri_kitap = os.path.join(self.paylasim, "ProjeOneri.xlsm")
        self.yonetim_kitap = os.path.join(self.yonetim, "ProjeYonetim.xlsm")

    def oneriler_yil(self, y=None):
        return os.path.join(self.oneriler, str(y or yil()))

    def degerlendirme_yil(self, y=None):
        return os.path.join(self.degerlendirme, str(y or yil()))

    def oneri_dosyalar(self):
        return _tara(self.oneriler, ".txt")

    def degerlendirme_dosyalar(self):
        return _tara(self.degerlendirme, ".txt")


def _tara(kok, uzanti):
    bulunan = []
    for dizin, _, dosyalar in os.walk(kok):
        for d in dosyalar:
            if d.lower().endswith(uzanti):
                bulunan.append(os.path.join(dizin, d))
    return sorted(bulunan)


def kayit_oku(yol):
    """Kayit dosyasini sozluk olarak dondurur (VBA tarafiyla ayni kurallar)."""
    with open(yol, "r", encoding="utf-8") as f:
        ham = f.read()
    sozluk = {}
    for satir in ham.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if "=" in satir:
            anahtar, _, deger = satir.partition("=")
            sozluk[anahtar.strip()] = deger.replace("<|>", "\n")
    return sozluk


def bom_var_mi(yol):
    with open(yol, "rb") as f:
        return f.read(3) == b"\xef\xbb\xbf"


@contextmanager
def ortam(yonetim_de=True):
    """OneDrive disinda gecici bir ortak klasor kurar, is bitince siler."""
    kok = tempfile.mkdtemp(prefix="proje_test_")
    o = Ortam(kok)
    try:
        os.makedirs(o.oneriler_yil(), exist_ok=True)
        os.makedirs(o.yonetim, exist_ok=True)
        os.makedirs(o.degerlendirme_yil(), exist_ok=True)
    
        kaynak_oneri = os.path.join(CIKTI, "ProjeOneri.xlsm")
        if not os.path.exists(kaynak_oneri):
            raise SystemExit("cikti\\ProjeOneri.xlsm yok. Önce: python kur.py")
        shutil.copy2(kaynak_oneri, o.oneri_kitap)

        if yonetim_de:
            kaynak_yonetim = os.path.join(CIKTI, "yonetim", "ProjeYonetim.xlsm")
            if not os.path.exists(kaynak_yonetim):
                raise SystemExit("cikti\\yonetim\\ProjeYonetim.xlsm yok. "
                                 "Önce: python kur.py")
            shutil.copy2(kaynak_yonetim, o.yonetim_kitap)

        yield o
    finally:
        shutil.rmtree(kok, ignore_errors=True)


@contextmanager
def excel(gorunur=False):
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = gorunur
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.ScreenUpdating = gorunur
    # Otomasyonla acilan dosyalarda makro uyarisi cikmasin.
    app.AutomationSecurity = MSO_OTOMASYON_DUSUK
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


@contextmanager
def kitap(app, yol, salt_okunur=False):
    """Kitabi acar.

    salt_okunur=True uretimdeki durumu taklit eder: ProjeOneri.xlsm tum
    personel icin salt okunurdur, boylece ilk acan dosyayi kilitlemez.
    Salt okunur acilan bir kitap bellekte duzenlenebilir (form doldurulabilir,
    makrolar ortak klasore yazabilir); yalnizca kitabin kendisi kaydedilemez.
    """
    wb = app.Workbooks.Open(os.path.abspath(yol), 0, salt_okunur)
    try:
        sessiz_mod(app, wb, True)
        yield wb
    finally:
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass


def sessiz_mod(app, wb, acik=True):
    """Kitaptaki mesaj kutularini kapatir.

    Gorunmez bir Excel'de acilan bir MsgBox ekranda gorunmez ama makro geri
    donmez: otomasyon sonsuza kadar bekler. Sessiz mod mesajlari ekrana
    cikarmak yerine kaydeder; testler hem kilitlenmez hem de kullaniciya ne
    soylendigini SonMesaj() ile dogrulayabilir.
    """
    try:
        app.Run(f"'{os.path.basename(wb.FullName)}'!modUI.SessizModAyarla", acik)
    except Exception:
        pass          # modUI icermeyen bir kitap olabilir


def derleme_sinamasi(app, wb):
    """ThisWorkbook modulunu derlenmeye zorlar. Derlenirse True doner.

    Belge modulundeki bir yordam adiyla dogrudan cagrilamaz; modul adiyla
    nitelenmesi gerekir ve bu ad Excel'in dil surumune gore yerellesir
    ("ThisWorkbook" / "BuÇalışmaKitabı"). Kitabin kendi CodeName'i kullanilir.

    Kanit donen degerde degil, cagrinin SORUNSUZ TAMAMLANMASINDADIR: modul
    derlenmiyorsa VBA bir hata penceresi acar; o da ya hata olarak doner ya da
    bekci tarafindan adiyla bildirilir.
    """
    try:
        calistir(app, wb, f"{wb.CodeName}.DerlemeSinamasi")
        return True
    except Exception:
        return False


def son_mesaj(app, wb):
    """Makronun kullaniciya gostermek istedigi son mesaji dondurur."""
    try:
        return app.Run(f"'{os.path.basename(wb.FullName)}'!modUI.SonMesaj") or ""
    except Exception:
        return ""


KORUMA_SIFRESI = "po-koruma"          # modAyar.SIFRE_KORUMA ile aynı


def korumasiz(ws):
    """Sayfa korumasini kaldirir.

    Testler giris alanlarini doldurmak icin korumayi acmak zorundadir.
    Bu is VBA'ya Application.Run ile bir Worksheet NESNESI gecirerek degil
    dogrudan COM uzerinden yapilir: nesne gecirmek pywin32'de guvenilmez.
    Makrolar korumayi zaten kendileri yonetir, testin geri kapatmasi gerekmez.
    """
    try:
        ws.Unprotect(KORUMA_SIFRESI)
    except Exception:
        pass                            # zaten korumasiz olabilir


# --------------------------------------------------------------------------
#  Bekci -- takilan makroyu adiyla bildirir
#
#  Gorunmez bir Excel'de acilan herhangi bir diyalog (VBA'nin calisma zamani
#  hata penceresi dahil) makroyu geri dondurmez ve testler sessizce sonsuza
#  kadar bekler. COM cagrisi iptal edilemez; bu yuzden bekci, hangi makronun
#  yanit vermedigini yazip sureci sonlandirir. Boylece takilma "zaman asimi"
#  degil, adi konmus bir hata olur.
# --------------------------------------------------------------------------
AZAMI_SURE = 90.0

_bekleyen = None
_bekci_basladi = False


def _bekci():
    while True:
        time.sleep(1.0)
        bekleyen = _bekleyen
        if bekleyen is None:
            continue
        makro, baslangic = bekleyen
        if time.monotonic() - baslangic > AZAMI_SURE:
            print(f"\n  !! Makro yanıt vermiyor: {makro}\n"
                  f"     ({AZAMI_SURE:.0f} sn boyunca döndü. Görünmez bir Excel'de "
                  f"açılmış bir\n      iletişim kutusu buna yol açar — çoğunlukla "
                  f"VBA'nın çalışma zamanı\n      hata penceresi.)",
                  flush=True)
            os._exit(2)


def calistir(app, wb, makro, *argumanlar):
    """Kitaba ait bir makroyu calistirir (dosya adiyla nitelenmis)."""
    global _bekleyen, _bekci_basladi

    if not _bekci_basladi:
        threading.Thread(target=_bekci, daemon=True).start()
        _bekci_basladi = True

    ad = f"'{os.path.basename(wb.FullName)}'!{makro}"
    _bekleyen = (makro, time.monotonic())
    try:
        return app.Run(ad, *argumanlar)
    finally:
        _bekleyen = None
