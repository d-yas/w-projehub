# -*- coding: utf-8 -*-
r"""Kurulum KONUMU testleri -- gerçek kullanıcı yolunu birebir izler.

NEDEN AYRI BIR TEST DOSYASI VAR
-------------------------------
Bir hata uretimde ortaya cikti ve butun testler gecerken gozden kacti:
"Seçiliyi Değerlendir" ile acilan bir oneri kaydedilirken hata veriyordu.
Sebebi kodda degil KONUMDAYDI -- kurulum OneDrive ile eslenen bir klasorde
duruyordu -- ve mevcut testlerin hicbiri bunu goremezdi, cunku:

  1) Hepsi %TEMP% altinda, yani OneDrive DISINDA kosuyor.
  2) Hepsi EnableEvents=False ile kosuyor, yani yonetim kitabinin
     Workbook_Open kodu (yedek + salt okunura gecis + uyari bandi) HIC
     calismiyor.

Bu dosya iki bosluu da kapatir: kitabi kullanicinin actigi gibi -- OLAYLAR
ACIK -- acar ve ayni akisi IKI FARKLI KONUMDA kosturur.

OLCULEN DAVRANIS
----------------
OneDrive ile eslenen bir klasordeki kitap bir Excel orneginde acik oldugu
surece -- SALT OKUNUR bile olsa -- ikinci bir Excel sureci onu yazma kipinde
acamiyor; Excel hata vermiyor, sessizce salt okunur aciyor. Ekip kitabi her
zaman acik oldugu icin boyle bir klasorde degerlendirme HIC kaydedilemez.
Ayni klasor OneDrive disindayken ayni islem sorunsuz calisiyor.

    python testler\test_konum.py
"""

import os
import shutil
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yardimci as y

import pythoncom
import win32com.client

# OneDrive konumunda kaydetme HEMEN reddedilmeli; yirmi saniye beklenmemeli.
HIZLI_RET_SN = 12.0


def _onedrive_koku():
    """Yerel bir OneDrive kokü (yoksa None).

    Ortam degiskenine ve kullanici profiline bakar -- modDosyaIO'daki VBA
    yordami da ayni iki yere bakar.
    """
    for ad in ("OneDrive", "OneDriveCommercial", "OneDriveConsumer"):
        deger = os.environ.get(ad, "")
        if deger and os.path.isdir(deger):
            return deger

    profil = os.environ.get("USERPROFILE", "")
    if profil and os.path.isdir(profil):
        for ad in sorted(os.listdir(profil)):
            if ad.lower().startswith("onedrive"):
                tam = os.path.join(profil, ad)
                if os.path.isdir(tam):
                    return tam
    return None


def _kurulum_kur(kok):
    """Verilen klasore temiz bir 'projeoneri\\' kurulumu koyar."""
    o = y.Ortam(kok)
    shutil.rmtree(o.paylasim, ignore_errors=True)
    os.makedirs(o.yonetim, exist_ok=True)
    for kaynak, hedef in (
        (os.path.join(y.CIKTI, "ProjeOneri.xlsm"), o.oneri_kitap),
        (os.path.join(y.CIKTI, "yonetim", y.DOSYA_YONETIM), o.yonetim_kitap),
    ):
        if not os.path.exists(kaynak):
            raise SystemExit(f"{kaynak} yok. Önce: python kur.py")
        shutil.copy2(kaynak, hedef)
    return o


def _gonderim_yap(o, baslik):
    """Personel tarafi: bir oneri gonderir, numarasini dondurur."""
    with y.excel() as app:
        with y.kitap(app, o.oneri_kitap, salt_okunur=True) as wb:
            no = y.calistir(app, wb, "modGonderim.TestGonderimi",
                            "Konum Testi", "5001",
                            "Konum testi için gönderim.", baslik,
                            "Bir çözüm önerisi.", "Fayda")
    del app
    return no


def _gercek_akis(o, no):
    r"""Kitabi KULLANICININ actigi gibi acar ve degerlendirmeyi kaydeder.

    Fark ince ama belirleyici: EnableEvents=True. Ancak o zaman
    Workbook_Open calisir -- yedek alinir, kitap salt okunura gecer ve
    Giris ekranindaki uyari bandi yazilir. Testlerin geri kalani bu yolu
    hic gormez.

    Sonuclari sozluk olarak dondurur.
    """
    sonuc = {}
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = True                 # <-- gercek kullanimdaki gibi
    app.ScreenUpdating = False
    wb = None
    try:
        wb = app.Workbooks.Open(os.path.abspath(o.yonetim_kitap), 0, False,
                                None, y.DOSYA_SIFRESI, "", True,
                                None, None, None, False)
        y.sessiz_mod(app, wb, True)

        sonuc["acilis_readonly"] = bool(wb.ReadOnly)
        sonuc["yedekler"] = o.yedek_dosyalar()
        sonuc["giris_bant"] = str(
            wb.Worksheets("Giriş").Range("giris_bant").Value or "")
        sonuc["onedrive"] = bool(
            y.calistir(app, wb, "modDepo.DepoOneDriveAltindaMi"))

        y.calistir(app, wb, "modKonsolide.SistemeGir")
        sonuc["sistemegir_mesaji"] = y.son_mesaj(app, wb)

        y.calistir(app, wb, "modDegerlendirme.OneriyiAc", no)
        sonuc["oneriyiac_mesaji"] = y.son_mesaj(app, wb)

        ws = wb.Worksheets("Değerlendirme")
        sonuc["ekrandaki_no"] = str(ws.Range("dg_oneri_no").Value or "")

        y.korumasiz(ws)
        ws.Range("dg_yeni_durum").Value = "Planlandı"
        ws.Range("dg_not").Value = "Konum testi kaydı."

        t0 = time.monotonic()
        y.calistir(app, wb, "modDegerlendirme.DegerlendirmeKaydet")
        sonuc["kaydet_suresi"] = time.monotonic() - t0
        sonuc["kaydet_mesaji"] = y.son_mesaj(app, wb)
        sonuc["kaydet_banti"] = str(ws.Range("dg_bant").Value or "")
    finally:
        try:
            if wb is not None:
                wb.Saved = True
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        del app
        pythoncom.CoUninitialize()

    sonuc["olaylar"] = y.depo_oku(o.yonetim_kitap)[1]
    y.okuyucuyu_kapat()
    return sonuc


# ==========================================================================
#  1. OneDrive DISINDA -- desteklenen kurulum, her sey calismali
# ==========================================================================
def _saglikli_konum(s):
    print("  · OneDrive dışında: gerçek akış baştan sona")

    kok = tempfile.mkdtemp(prefix="proje_konum_")
    try:
        o = _kurulum_kur(kok)
        no = _gonderim_yap(o, "OneDrive dışı öneri")
        s.kontrol("Gönderim numara aldı", str(no).startswith("PRJ-"), str(no))
        if not str(no).startswith("PRJ-"):
            return

        r = _gercek_akis(o, no)

        s.kontrol("Konum OneDrive altında değil", not r["onedrive"])
        s.kontrol("Açılışta kitap salt okunura geçti", r["acilis_readonly"])
        s.esit("Açılışta günlük yedek alındı", 1, len(r["yedekler"]))
        s.esit("Giriş ekranında uyarı bandı yok", "", r["giris_bant"])
        s.esit("Sisteme giriş hatasız", "", r["sistemegir_mesaji"])
        s.esit("Öneri ekrana hatasız yüklendi", "", r["oneriyiac_mesaji"])
        s.esit("Doğru öneri ekranda", no, r["ekrandaki_no"])

        s.esit("Değerlendirme hatasız kaydedildi", "", r["kaydet_mesaji"])
        s.kontrol("Onay bandı yazıldı", "kaydedildi" in r["kaydet_banti"],
                  r["kaydet_banti"])
        s.esit("Olay diske yazıldı", 1, len(r["olaylar"]))
        if r["olaylar"]:
            s.esit("Olay doğru öneriye bağlandı", no, r["olaylar"][0]["oneri_no"])
            s.esit("Olayda yeni durum var", "Planlandı",
                   r["olaylar"][0]["yeni_durum"])
        s.kontrol("Kaydetme makul sürede bitti (< 30 sn)",
                  r["kaydet_suresi"] < 30.0, f"{r['kaydet_suresi']:.1f} sn")
        s.kontrol("Geriye kilit dosyası kalmadı",
                  not o.sahiplik_dosyalari(), str(o.sahiplik_dosyalari()))
    finally:
        shutil.rmtree(kok, ignore_errors=True)


# ==========================================================================
#  2. OneDrive ALTINDA -- desteklenmeyen kurulum, ADIYLA reddedilmeli
#
#  Burada aranan sey "calissin" degil. Orada calismasi Excel'in elinde
#  degil. Aranan sey, sistemin durumu ANINDA ve DOGRU sebeple soylemesi:
#  kullanici yirmi saniye beklememeli ve "başka biri yazıyor" gibi yanlis
#  bir sebep okumamali -- kimse yazmiyor, konum yanlis.
# ==========================================================================
def _onedrive_konumu(s):
    kok_od = _onedrive_koku()
    if not kok_od:
        print("  · OneDrive klasörü yok — konum reddi testi atlanıyor")
        return

    print("  · OneDrive altında: kaydetme adıyla reddedilmeli")
    kok = os.path.join(kok_od, "_proje_konum_testi")
    try:
        o = _kurulum_kur(kok)

        # Gonderim OneDrive altinda da CALISIR: personelin Excel'inde yonetim
        # kitabi acik degildir. Engellenmemeli.
        no = _gonderim_yap(o, "OneDrive öneri")
        s.kontrol("Gönderim OneDrive altında da çalışıyor",
                  str(no).startswith("PRJ-"), str(no))
        if not str(no).startswith("PRJ-"):
            return

        r = _gercek_akis(o, no)

        s.kontrol("Konum OneDrive altında olarak tanındı", r["onedrive"])
        s.kontrol("Giriş ekranı açılışta uyarıyor",
                  "OneDrive" in r["giris_bant"], repr(r["giris_bant"]))

        # Okuma tarafi calismali: yalnizca YAZMA imkansiz.
        s.esit("Sisteme giriş yine de çalışıyor", "", r["sistemegir_mesaji"])
        s.esit("Öneri ekrana yine de yükleniyor", "", r["oneriyiac_mesaji"])
        s.esit("Doğru öneri ekranda", no, r["ekrandaki_no"])

        mesaj = r["kaydet_mesaji"]
        s.kontrol("Kaydetme reddedildi", mesaj.startswith("hata:"), mesaj[:120])
        s.kontrol("Ret gerekçesi OneDrive olarak veriliyor",
                  "OneDrive" in mesaj, mesaj[:200])
        s.kontrol("Yanıltıcı “başka biri yazıyor” gerekçesi verilmiyor",
                  "başka biri yazıyor" not in mesaj
                  and "başka bir kullanıcı" not in mesaj, mesaj[:200])
        s.kontrol(f"Ret HEMEN geliyor (< {HIZLI_RET_SN:.0f} sn)",
                  r["kaydet_suresi"] < HIZLI_RET_SN,
                  f"{r['kaydet_suresi']:.1f} sn")
        s.esit("Onay bandı yazılmadı", "", r["kaydet_banti"])
        s.esit("Diske hiçbir olay yazılmadı", 0, len(r["olaylar"]))
        s.kontrol("Geriye kilit dosyası kalmadı",
                  not o.sahiplik_dosyalari(), str(o.sahiplik_dosyalari()))
    finally:
        shutil.rmtree(os.path.join(kok, "projeoneri"), ignore_errors=True)
        shutil.rmtree(kok, ignore_errors=True)


def calistir():
    s = y.Sonuc("Konum")
    baslangic_excel = y.excel_sayisi()

    _saglikli_konum(s)
    _onedrive_konumu(s)

    kalan = y.excel_sayisi_bekle(baslangic_excel)
    s.kontrol("Arkada görünmez Excel süreci kalmadı", kalan <= baslangic_excel,
              f"(başlangıç: {baslangic_excel}, bitiş: {kalan})")

    return s.bitir()


if __name__ == "__main__":
    sys.exit(calistir())
