# -*- coding: utf-8 -*-
r"""NTFS izinlerini gercek izinlerle dogrular.

BU SURUMDE IZIN MODELI DEGISTI. Eskiden "yonetim\oneriler\" bir BIRAKMA
KUTUSUYDU: personel oraya yazabilir ama iceriini goremezdi. Veri artik ayri
dosyalarda degil, yonetim kitabinin ICINDE. Excel bir dosyayi okumadan
yazamaz; dolayisiyla personelin o dosya uzerinde okuma hakki da olmak
ZORUNDA. Bunun uc sonucu var ve ucu de bilincli kabul edilmistir:

  1) Gizliligin siniri artik NTFS degil, dosyanin ACILIS PAROLASIDIR.
     Personel dosyayi kopyalayabilir ama parolasiz acamaz. Tehdit modeli
     siradan personeldir; parolayi VBA'dan cikarabilecek biri icin bu bir
     sinir degildir (bkz. TASARIM-VE-GEREKCE.md madde 8).

  2) Izinler DOSYAYA degil KLASORE, mirasla verilir. Excel kaydederken
     dosyayi yerinde degistirmez; gecici bir dosya yazip aslinin yerine
     koyar. Dosyaya verilen acik haklar bu sirada kaybolabilir, klasorden
     miras alinanlar kalir. Bu test kaydetmenin ardindan dosyanin hala
     erisilebilir oldugunu dogrular.

  3) Personelin klasorde dosya olusturma ve silme hakki olmak zorundadir
     (Excel'in gecici dosyasi ve "~$" sahiplik dosyasi icin). Korunmasi
     gereken tek yer YEDEK KLASORUDUR: veri kaybina karsi son siginak orasi
     oldugu icin personele kapali olmalidir.

    python testler\test_izinler.py

IZIN UYGULAMA SIRASI ONEMLIDIR: once alt klasor (yedek), sonra ust klasor.
Ters sirada yapilirsa izin komutlarinin kendisi calisamaz hale gelir.

NOT: Test kendi kullanicisi uzerinde calisir. Bir kullanicinin izinlerini
kisitlayip kendi kodumuzu o kisitlar altinda kosturmak, iki ayri hesap
gerektirmeden ulasilabilecek en yakin gercekliktir. icacls kullanilamiyorsa
(ornegin FAT32 bir diskte) test atlanir.
"""

import os
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yardimci as y

KULLANICI = f"{os.environ.get('USERDOMAIN', '')}\\{os.environ.get('USERNAME', '')}"

# Personelin "yonetim\" klasorunde ihtiyaci olan haklar: Degistir.
#   Okuma + yazma + dosya olusturma + SILME (Excel'in gecici dosyasi icin).
# Daha azi yetmez: gonderim, kitabi kaydetmek demektir.
YONETIM_HAKLARI = "(OI)(CI)(M)"

# Yedek klasoru: personele hicbir hak yok. Yalnizca "izinleri gor" birakilir
# ki test sonunda klasoru geri acabilsin.
YEDEK_HAKLARI = "(RC)"


def icacls(*argumanlar):
    sonuc = subprocess.run(["icacls", *argumanlar], capture_output=True, text=True)
    return sonuc.returncode, ((sonuc.stdout or "") + (sonuc.stderr or "")).strip()


def _engellendi_mi(islem):
    """islem izin hatasi verdiyse True."""
    try:
        islem()
        return False
    except OSError:
        return True


def calistir():
    s = y.Sonuc("İzinler")
    baslangic_excel = y.excel_sayisi()

    kok = tempfile.mkdtemp(prefix="proje_izin_")
    o = y.Ortam(kok)

    try:
        os.makedirs(o.yonetim, exist_ok=True)
        os.makedirs(o.yedek, exist_ok=True)

        for kaynak, hedef in (
            (os.path.join(y.CIKTI, "ProjeOneri.xlsm"), o.oneri_kitap),
            (os.path.join(y.CIKTI, "yonetim", y.DOSYA_YONETIM), o.yonetim_kitap),
        ):
            if not os.path.exists(kaynak):
                raise SystemExit(f"{kaynak} yok. Önce: python kur.py")
            shutil.copy2(kaynak, hedef)

        # Yedek klasorunde personelin gormemesi gereken bir kopya
        gizli_yedek = os.path.join(o.yedek, "ProjeYonetim_20260101.xlsm")
        shutil.copy2(o.yonetim_kitap, gizli_yedek)

        # --- 1) ÖNCE alt klasör: yedek --------------------------------------
        print("  · yedek klasörü personele kapatılıyor")
        kod, cikti = icacls(o.yedek, "/inheritance:r")
        if kod != 0:
            print(f"\n  icacls kullanılamıyor, test atlanıyor:\n  {cikti}")
            return 0
        kod, cikti = icacls(o.yedek, "/grant", f"{KULLANICI}:{YEDEK_HAKLARI}")
        if kod != 0:
            print(f"\n  izin verilemedi, test atlanıyor:\n  {cikti}")
            return 0

        # --- 2) SONRA üst klasör: yonetim -----------------------------------
        icacls(o.yonetim, "/inheritance:r")
        icacls(o.yonetim, "/grant", f"{KULLANICI}:{YONETIM_HAKLARI}")

        # --- Kilitliyken neler engellenmiş? ---------------------------------
        s.kontrol("Yedek klasörü listelenemiyor",
                  _engellendi_mi(lambda: os.listdir(o.yedek)))
        s.kontrol("Yedek kopya okunamıyor",
                  _engellendi_mi(lambda: open(gizli_yedek, "rb").read()))
        s.kontrol("Yedek klasörüne yazılamıyor",
                  _engellendi_mi(
                      lambda: open(os.path.join(o.yedek, "x.txt"), "w").write("x")))

        # Yonetim klasoru personele ACIKTIR -- bu bilincli bir kabuldur.
        # Gizliligi saglayan sey klasor izni degil, dosyanin parolasidir.
        s.kontrol("Yönetim klasörü personele açık (bilinçli kabul)",
                  not _engellendi_mi(lambda: os.listdir(o.yonetim)))
        s.kontrol("Depo dosyası şifreli (gizliliğin gerçek sınırı)",
                  open(o.yonetim_kitap, "rb").read(8)
                  == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")

        onceki_izin = subprocess.run(["icacls", o.yonetim_kitap],
                                     capture_output=True, text=True).stdout

        # --- Kilitliyken gerçek gönderim makrosu ---------------------------
        print("  · kısıtlı izinlerle iki gönderim")
        numaralar = []
        with y.excel() as app:
            with y.kitap(app, o.oneri_kitap, salt_okunur=True) as wb:
                s.esit("Kök klasör kısıtlı yapıda da bulunuyor",
                       os.path.normcase(o.paylasim),
                       os.path.normcase(y.calistir(app, wb, "modAyar.KokKlasor")))
                for i in (1, 2):
                    numaralar.append(y.calistir(
                        app, wb, "modGonderim.TestGonderimi",
                        f"Kullanıcı {i}", f"1000{i}",
                        "Kısıtlı izinlerle yazma denemesi.", f"Öneri {i}",
                        "Bir çözüm önerisi.", "Fayda"))

            s.kontrol("Değiştir hakkıyla gönderim yapılabildi",
                      all(str(n).startswith("PRJ-") for n in numaralar),
                      str(numaralar))
            s.esit("İki gönderim iki farklı numara aldı", 2, len(set(numaralar)))

            # Kaydetme dosyayi yerine koyar; klasorden miras alinan haklarin
            # bunu atlatmasi gerekir. Aksi halde ILK gonderimden sonra dosya
            # erisilemez olur ve sistem sessizce durur.
            s.kontrol("Kaydetmeden sonra depo hâlâ okunabiliyor",
                      not _engellendi_mi(lambda: open(o.yonetim_kitap, "rb").read()))
            sonraki_izin = subprocess.run(["icacls", o.yonetim_kitap],
                                          capture_output=True, text=True).stdout
            s.kontrol("Kaydetmeden sonra dosyanın izinleri korundu",
                      onceki_izin.strip() == sonraki_izin.strip(),
                      f"önce: {onceki_izin.strip()[:120]} / "
                      f"sonra: {sonraki_izin.strip()[:120]}")

            # Yedek alinamiyor ama HATA DA VERMIYOR: yedek alamamak kitabin
            # acilmasini engellememelidir.
            with y.kitap(app, o.yonetim_kitap, salt_okunur=True) as wb:
                y.calistir(app, wb, "modDepo.YedekAl")
                s.esit("Yedek alınamayınca kullanıcıya hata gösterilmiyor",
                       "", y.son_mesaj(app, wb))

            y.okuyucuyu_kapat()

        # "with ... as app" blogu bitince degisken BAGLI KALIR ve Excel
        # sureci, kendisine ait son COM vekili birakilana kadar olmez.
        del app

        # --- İzinleri aç, sonuca bak ----------------------------------------
        icacls(o.yonetim, "/reset", "/t", "/c", "/q")
        icacls(o.yonetim, "/grant", f"{KULLANICI}:(OI)(CI)(F)", "/t", "/c", "/q")

        s.esit("Yedek klasörüne yeni dosya yazılmamış",
               ["ProjeYonetim_20260101.xlsm"], o.yedek_dosyalar())

        oneriler, _ = y.depo_oku(o.yonetim_kitap)
        y.okuyucuyu_kapat()
        s.esit("Her iki öneri de gerçekten yazılmış", 2, len(oneriler))
        s.esit("Depodaki numaralar döndürülenlerle aynı",
               sorted(numaralar), sorted(r["oneri_no"] for r in oneriler))
        s.kontrol("Kayıtlar eksiksiz ve okunabilir",
                  all(r["sema"] and r["tarih"]
                      and "Kısıtlı izinlerle" in r["mevcut_durum"]
                      for r in oneriler),
                  str(oneriler[:1]))

    finally:
        icacls(o.yonetim, "/reset", "/t", "/c", "/q")
        icacls(o.yonetim, "/grant", f"{KULLANICI}:(OI)(CI)(F)", "/t", "/c", "/q")
        shutil.rmtree(kok, ignore_errors=True)

    # Aranan sey "sifir Excel" degil, TESTIN sizdirmadigidir: baslangicta
    # baska bir isten kalan bir ornek kapanmak uzere olabilir.
    kalan = y.excel_sayisi_bekle(baslangic_excel)
    s.kontrol("Arkada görünmez Excel süreci kalmadı", kalan <= baslangic_excel,
              f"(başlangıç: {baslangic_excel}, bitiş: {kalan})")

    return s.bitir()


if __name__ == "__main__":
    sys.exit(calistir())
