# -*- coding: utf-8 -*-
r"""Test altyapisi: sahte ortak klasor + gercek Excel oturumu.

ONEMLI -- testler neden gecici klasorde calisir:
Proje klasoru OneDrive ile eslenmis olabilir. Eslenmis klasorlerde Excel
kitabin konumunu disk yolu yerine "https://..." adresi olarak bildirir ve
Guvenilir Konum ayari beklendigi gibi davranmaz. Testler bu belirsizligi
disarida birakmak icin %TEMP% altinda -- OneDrive DISINDA -- kendi
"ortak klasorunu" kurar. Uretim ortami zaten bir UNC paylasimi olacaktir.

ONEMLI -- parola her acista verilmelidir:
Yonetim kitabi acilis parolasiyla sifrelidir. Parolasi verilmeden acilirsa
gorunmez Excel bir parola diyalogu acar; diyalog ekranda gorunmez ama cagri
GERI DONMEZ. Bu olculdu (40 sn sonra hala bekliyordu). Bu yuzden kitabi acan
her yol parolayi acikca gecer.
"""

import gc
import os
import shutil
import subprocess
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

sys.path.insert(0, KOK)
import kur  # noqa: E402  -- sabitleri modAyar.bas'tan okumak icin

# Cikti bir boruya yonlendirildiginde Python yerel kod sayfasini kullanir ve
# testlerin yazdigi "·" gibi karakterler UnicodeEncodeError verir.
for _akis in (sys.stdout, sys.stderr):
    try:
        _akis.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

MSO_OTOMASYON_DUSUK = 1        # msoAutomationSecurityLow
XL_YUKARI = -4162              # xlUp
XL_SOLA = -4159                # xlToLeft

# Sifreler ve depo bilgisi TEK kaynaktan -- modAyar.bas -- okunur.
DOSYA_SIFRESI = kur.ayar_sabiti("SIFRE_DOSYA")
KORUMA_SIFRESI = kur.ayar_sabiti("SIFRE_KORUMA")
DOSYA_YONETIM = kur.ayar_sabiti("DOSYA_YONETIM")
KLASOR_YEDEK = kur.ayar_sabiti("KLASOR_YEDEK")
YEDEK_ADET = kur.ayar_sayisi("YEDEK_ADET")

SAYFA_ONERILER = "Oneriler"
SAYFA_OLAYLAR = "Olaylar"

# Depoyu okurken kac kez denenecegi. Cakisma altinda salt okunur bir
# acilis bile basarisiz olabilir: dosya o anda bir yazicinin "gecici
# dosyayi aslinin yerine koyma" adiminda olabilir.
AZAMI_OKUMA_DENEMESI = 15


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
    r"""Gecici bir 'projeoneri\' paylasim klasoru ve icindeki iki kitap.

    Ayri kayit dosyasi ve yil klasoru YOKTUR: butun veri yonetim kitabinin
    icindedir. Klasorde uretilen tek diger sey gunluk yedek kopyalardir.
    """

    def __init__(self, kok):
        self.kok = kok
        self.paylasim = os.path.join(kok, "projeoneri")
        self.yonetim = os.path.join(self.paylasim, "yonetim")
        self.yedek = os.path.join(self.yonetim, KLASOR_YEDEK)
        self.oneri_kitap = os.path.join(self.paylasim, "ProjeOneri.xlsm")
        self.yonetim_kitap = os.path.join(self.yonetim, DOSYA_YONETIM)

    def yedek_dosyalar(self):
        if not os.path.isdir(self.yedek):
            return []
        return sorted(d for d in os.listdir(self.yedek) if d.lower().endswith(".xlsm"))

    def sahiplik_dosyalari(self):
        """Geride kalmis kilit dosyalari.

        Iki tur vardir ve ikisi de islem bitince YOK OLMALIDIR:
        Excel'in kendi "~$" sahiplik dosyasi ve sistemin yazma kilidi
        (".kilit"). Kalan bir kilit sonraki yazicilari iki dakika boyunca
        bekletir.
        """
        bulunan = []
        for dizin, _, dosyalar in os.walk(self.paylasim):
            bulunan += [d for d in dosyalar
                        if d.startswith("~$") or d.endswith(".kilit")]
        return bulunan


@contextmanager
def ortam():
    """OneDrive disinda gecici bir ortak klasor kurar, is bitince siler."""
    kok = tempfile.mkdtemp(prefix="proje_test_")
    o = Ortam(kok)
    try:
        os.makedirs(o.yonetim, exist_ok=True)

        for kaynak, hedef in (
            (os.path.join(CIKTI, "ProjeOneri.xlsm"), o.oneri_kitap),
            (os.path.join(CIKTI, "yonetim", DOSYA_YONETIM), o.yonetim_kitap),
        ):
            if not os.path.exists(kaynak):
                raise SystemExit(f"{kaynak} yok. Önce: python kur.py")
            shutil.copy2(kaynak, hedef)

        yield o
    finally:
        shutil.rmtree(kok, ignore_errors=True)


# CoInitialize/CoUninitialize YALNIZCA en distaki excel() cagrisinda yapilir.
#
# Bir test ikinci bir Excel ornegine ihtiyac duyar (ornegin depoyu kilitli
# tutan bir "tutucu"). Ic ice iki excel() blogunda ictekinin cikisi
# CoUninitialize cagirirsa is parcaciginin COM apartmani kapanir ve DISTAKI
# uygulamanin butun nesneleri "sunucuya bagli degil" hatasi verir -- oysa o
# Excel hala calisiyordur. (Bu hata boyle bulundu.)
_com_derinlik = 0


def _com_ac():
    global _com_derinlik
    if _com_derinlik == 0:
        pythoncom.CoInitialize()
    _com_derinlik += 1


def _com_kapa():
    global _com_derinlik
    _com_derinlik -= 1
    if _com_derinlik == 0:
        pythoncom.CoUninitialize()


@contextmanager
def excel(gorunur=False):
    _com_ac()

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
        # Excel sureci ancak KENDISINE ait butun COM vekilleri birakildiginda
        # olur. Python'da bir Worksheet/Range vekili hala referanslanmissa Quit
        # cagrisi surece dokunmaz. Tek bir testte bu gorunmez -- surec bitince
        # her sey birakilir -- ama testler ARDISIK kostugunda onceki asamadan
        # kalan vekiller Excel'i ayakta tutar ve "sizinti" gibi gorunur.
        gc.collect()
        _com_kapa()


def sifresi(yol):
    """Yonetim kitabi parolalidir, oneri kitabi degildir."""
    if os.path.basename(yol).lower() == DOSYA_YONETIM.lower():
        return DOSYA_SIFRESI
    return ""


def _ac_ham(app, yol, salt_okunur):
    """Kitabi URETIMDEKI ile ayni acis ayarlariyla acar.

    Notify:=False kritiktir. Varsayilan (True) ile, dosya o anda baskasinda
    aciksa Excel dosyayi bildirim kuyruguna ekleyip HATA VERIR; salt okunur
    acmayi bile denemez. modDepo.DepoAc bu yuzden Notify:=False geciyor ve
    testler de gecmeli -- yoksa testler uretimden farkli davranan bir yolu
    sinar ve cakisma altinda kendileri patlar.

    Konumsal sira: FileName, UpdateLinks, ReadOnly, Format, Password,
    WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter,
    Editable, Notify
    """
    return app.Workbooks.Open(os.path.abspath(yol), 0, salt_okunur, None,
                              sifresi(yol), "", True, None, None, None, False)


@contextmanager
def kitap(app, yol, salt_okunur=False):
    """Kitabi acar.

    salt_okunur=True uretimdeki durumu taklit eder:

    · ProjeOneri.xlsm tum personel icin salt okunurdur, boylece ilk acan
      dosyayi kilitlemez.
    · ProjeYonetim.xlsm'i ekip normalde yazma kipinde acar ama Workbook_Open
      ANINDA salt okunura gecer (modDepo.SaltOkunuraGec), yoksa gun boyu
      kilidi tutar ve personel oneri gonderemez. Testler EnableEvents=False
      ile calistigi icin Workbook_Open CALISMAZ; ekibi taklit eden her test bu
      yuzden kitabi acikca salt_okunur=True ile acmalidir.
    """
    wb = _ac_ham(app, yol, salt_okunur)
    try:
        sessiz_mod(app, wb, True)
        yield wb
    finally:
        try:
            wb.Saved = True
            wb.Close(SaveChanges=False)
        except Exception:
            pass


def ac(app, yol, salt_okunur=True):
    """Kitabi acar ve wb dondurur (baglam yoneticisi olmadan).

    Parola her zaman gecilir; eksik parola gorunmez Excel'i kilitler.
    Yonetim kitabi VARSAYILAN OLARAK salt okunur acilir: yazma kilidini
    tutmak, depoya yazmasi gereken gizli ornegi otuz saniye bekletir.
    """
    wb = _ac_ham(app, yol, salt_okunur)
    sessiz_mod(app, wb, True)
    return wb


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
#  DEPOYU OKUMA -- diskteki gercek veri
# --------------------------------------------------------------------------
def _sayfa_oku(wb, ad):
    """Bir depo sayfasini sozluk listesi olarak dondurur."""
    ws = wb.Worksheets(ad)
    son_sutun = ws.Cells(1, ws.Columns.Count).End(XL_SOLA).Column
    basliklar = [str(ws.Cells(1, c).Value or "") for c in range(1, son_sutun + 1)]

    son = ws.Cells(ws.Rows.Count, 1).End(XL_YUKARI).Row
    if son < 2:
        return []

    veri = ws.Range(ws.Cells(2, 1), ws.Cells(son, son_sutun)).Value
    satirlar = []
    for satir in veri:
        satirlar.append({b: ("" if d is None else str(d))
                         for b, d in zip(basliklar, satir)})
    return satirlar


# --------------------------------------------------------------------------
#  Okuyucu ornegi -- depoyu testin kendi Excel'inden AYRI bir surecte okur.
#
#  Neden ayri: Excel ayni dosyayi ayni ornekte iki kez acamaz. Testin kitabi
#  aciyken ayni yolu tekrar Open etmek YENI bir kitap vermez, ZATEN ACIK olani
#  verir; sonraki Close o kitabi kapatir ve testin elindeki nesne "istemciden
#  ayrilmis" olur. (Bu hata boyle bulundu.)
#
#  Cagri basina yeni bir Excel baslatmak pahali oldugu icin tek bir okuyucu
#  ornegi acilir ve surecin sonunda kapatilir.
# --------------------------------------------------------------------------
_okuyucu_app = None


def _okuyucu():
    global _okuyucu_app
    if _okuyucu_app is None:
        # Okuyucu bir excel() blogunun DISINDAN da cagrilabilir; COM'u kendisi
        # baslatir ve okuyucuyu_kapat() ile birakir.
        _com_ac()
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.ScreenUpdating = False
        app.AutomationSecurity = MSO_OTOMASYON_DUSUK
        _okuyucu_app = app
    return _okuyucu_app


def _okuyucuyu_unut():
    """Olu okuyucu vekilini atar; COM sayaci bozulmaz.

    Surecin kendisi zaten kapanmis oldugu icin Quit denenmez -- denenirse
    yeni bir hata verir ve asil hatayi gizler.
    """
    global _okuyucu_app
    if _okuyucu_app is not None:
        _okuyucu_app = None
        gc.collect()
        _com_kapa()


def okuyucuyu_kapat():
    """Okuyucu ornegini kapatir. EXCEL.EXE sayisi karsilastirilmadan ONCE.

    SIRA ONEMLIDIR. Quit'ten hemen sonra CoUninitialize cagirmak, Excel
    kapanisini bitirmeden apartmani yikar ve surec ayakta kalir. Once vekil
    birakilir, cikis icin kisa bir sure taninir, COM apartmani EN SON kapanir.
    (Ardisik kosan testlerde bir EXCEL.EXE'nin dakikalarca kalmasinin sebebi
    buydu; tek basina kosan testte gorunmuyordu, cunku orada okuyucu bir
    excel() blogunun icinde kapatiliyor ve CoUninitialize gecikiyordu.)
    """
    global _okuyucu_app
    if _okuyucu_app is not None:
        try:
            _okuyucu_app.Quit()
        except Exception:
            pass
        _okuyucu_app = None
        gc.collect()
        time.sleep(2.0)
        _com_kapa()


def depo_oku(yol):
    """(oneriler, olaylar) -- DISKTEKI depo satirlari.

    SALT OKUNUR acilir: bir test okurken baska bir surecin yazmasini
    engellememelidir.

    Okuyucu ornegi KENDINI ONARIR. Otomasyonla acilan bir Excel, son kitabi
    da kapandiktan sonra elinde is kalmadigina karar verip kapanabilir; bu
    ornek okumalar arasinda bos kaldigi icin (paralel bolumde dakikalarca)
    tam olarak bu oluyor ve bir sonraki cagri "Arabirim bilinmiyor" hatasi
    veriyordu. Olu vekil atilir, yenisi acilir.
    """
    son_hata = None
    for deneme in range(1, AZAMI_OKUMA_DENEMESI + 1):
        app = _okuyucu()
        try:
            wb = _ac_ham(app, yol, True)
        except Exception as hata:
            # Iki ayri sebep, ayni belirti: (a) vekil olmus, (b) dosya o anda
            # bir yazicinin "gecici dosyayi yerine koyma" adiminda. Ikisi de
            # gecicidir; okuyucu atilir ve kisa bir bekleyisten sonra yeniden
            # denenir. Uretim tarafinda ayni isi modDepo.DepoAc yapar.
            son_hata = hata
            _okuyucuyu_unut()
            time.sleep(1.0)
            continue
        try:
            return _sayfa_oku(wb, SAYFA_ONERILER), _sayfa_oku(wb, SAYFA_OLAYLAR)
        finally:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass

    raise RuntimeError(
        f"Depo {AZAMI_OKUMA_DENEMESI} denemede okunamadı: {yol}") from son_hata


def excel_sayisi():
    """Calisan EXCEL.EXE sayisi.

    modDepo her yazma icin gizli bir Excel ornegi acar ve kapatir. Kapatilmayan
    bir ornek arkada dosya kilidini tutmaya devam eder; testler bu yuzden
    baslangic ve bitis sayilarini karsilastirir.
    """
    cikti = subprocess.run(["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE"],
                           capture_output=True, text=True).stdout
    return 0 if "No tasks" in cikti else cikti.count("EXCEL.EXE")


def excel_sayisi_bekle(hedef, saniye=90.0):
    """Sayi hedefe dusene kadar bekler, son olculen degeri dondurur.

    Quit cagrisi ANINDA sonuc vermez: Excel'e "kapan" demek ile surecin
    gercekten olmesi arasinda onlarca saniye gecebilir -- ozellikle bir
    testte arka arkaya kirktan fazla gizli ornek acilip kapandiktan sonra.
    Aranan sey gecikme degil SIZINTIDIR: kapanmayan bir ornek dosya kilidini
    SURESIZ tutar. Bu yuzden genis ama sonlu bir pencere beklenir; bes
    saniyeyi asan bekleme ayrica yazilir ki yavaslamak sessizce normallesmesin.
    """
    gc.collect()                 # birakilmamis COM vekilleri Excel'i yasatir
    t0 = time.monotonic()
    bitis = t0 + saniye
    sayi = excel_sayisi()
    while sayi > hedef and time.monotonic() < bitis:
        time.sleep(1.0)
        gc.collect()
        sayi = excel_sayisi()
    gecen = time.monotonic() - t0
    if gecen > 5.0:
        print(f"    (Excel süreçlerinin kapanması {gecen:.0f} sn sürdü)",
              flush=True)
    return sayi


# --------------------------------------------------------------------------
#  Bekci -- takilan makroyu adiyla bildirir
#
#  Gorunmez bir Excel'de acilan herhangi bir diyalog (VBA'nin calisma zamani
#  hata penceresi dahil) makroyu geri dondurmez ve testler sessizce sonsuza
#  kadar bekler. COM cagrisi iptal edilemez; bu yuzden bekci, hangi makronun
#  yanit vermedigini yazip sureci sonlandirir. Boylece takilma "zaman asimi"
#  degil, adi konmus bir hata olur.
#
#  Sure 90'dan 180'e cikarildi: her yazma artik gizli bir Excel ornegi
#  baslatiyor (~1 sn) ve kilit catismasinda 30 sn'ye kadar yeniden deniyor.
#  Bekci gecikmeyi degil, GERCEK kilitlenmeyi yakalamali.
# --------------------------------------------------------------------------
AZAMI_SURE = 180.0

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
                  f"VBA'nın çalışma zamanı\n      hata penceresi ya da eksik "
                  f"parola.)\n"
                  f"     Arkada görünmez Excel süreçleri kalmış olabilir:\n"
                  f"       taskkill /F /IM EXCEL.EXE",
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
