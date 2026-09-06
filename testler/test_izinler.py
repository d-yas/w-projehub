# -*- coding: utf-8 -*-
r"""Birakma kutusu izinlerini gercekten dogrular.

Sistemin gizlilik iddiasi sudur: personel "yonetim\oneriler\" klasorune oneri
BIRAKABILIR ama yonetim tarafinda hicbir seyi GOREMEZ -- ne klasorleri
listeleyebilir, ne baskasinin onerisini okuyabilir, ne degerlendirme notlarina
ulasabilir, ne de bir dosyayi silebilir.

Bu test o iddiayi gercek NTFS izinleriyle sinar: klasorleri kilitler, gercek
gonderim makrosunu calistirir, sonra izinleri geri acip sonuca bakar.

    python testler\test_izinler.py

IZIN UYGULAMA SIRASI ONEMLIDIR (test bunu da ornekler):
once alt klasorler ayarlanir, "yonetim\" EN SON kisitlanir. Ters sirada
yapilirsa izin komutlarinin kendisi calisamaz hale gelir ve klasorler
sessizce erisilemez kalir.

NOT: Test kendi kullanicisi uzerinde calisir. Bir kullanicinin izinlerini
kisitlayip kendi kodumuzu o kisitlar altinda kosturmak, iki ayri hesap
gerektirmeden ulasilabilecek en yakin gerceklige denk gelir. icacls
kullanilamiyorsa (ornegin FAT32 bir diskte) test atlanir.
"""

import os
import shutil
import subprocess
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yardimci as y

KULLANICI = f"{os.environ.get('USERDOMAIN', '')}\\{os.environ.get('USERNAME', '')}"

# Personel grubuna "yonetim\oneriler\" uzerinde verilecek haklar.
#   WD  dosya olustur / veri yaz      AD  klasor olustur / veri ekle
#   X   klasorde gezin                RA  oznitelik oku      REA  gen. oznitelik oku
#   WA  oznitelik yaz                 WEA gen. oznitelik yaz  RC  izinleri oku
# BILINCLI OLARAK YOK:  RD (listele / oku),  DE (sil),  DC (alt oge sil)
BIRAKMA_KUTUSU = "(OI)(CI)(WD,AD,X,RA,REA,WA,WEA,RC)"

# "yonetim\" klasoru: YALNIZCA GECIS, miras bayragi YOK.
# Miras bayragi olmadigi icin kardes klasore (degerlendirme) hicbir hak
# sizmaz.
YONETIM_GECIS = "(X)"


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

    kok = tempfile.mkdtemp(prefix="proje_izin_")
    o = y.Ortam(kok)
    yil_klasoru = o.oneriler_yil()

    try:
        # Yil klasoru BILEREK acilmaz: uretimde de kodun kendisi olusturur ve
        # izinleri "oneriler"den miras alir.
        for d in (o.oneriler, o.degerlendirme):
            os.makedirs(d, exist_ok=True)

        kaynak = os.path.join(y.CIKTI, "ProjeOneri.xlsm")
        if not os.path.exists(kaynak):
            raise SystemExit("cikti\\ProjeOneri.xlsm yok. Önce: python kur.py")
        shutil.copy2(kaynak, o.oneri_kitap)

        # Personelin asla gormemesi gereken bir degerlendirme notu
        gizli_not = os.path.join(o.degerlendirme, "PRJ-26AAA_2026010100000000_1A2B.txt")
        with open(gizli_not, "w", encoding="utf-8") as f:
            f.write("oneri_no=PRJ-26AAA\nkarar_notu=Gizli karar gerekçesi\n"
                    "kayit_sonu=1\n")
        with open(o.yonetim_kitap, "wb") as f:
            f.write(b"yonetim kitabinin yerine gecen dosya")

        # --- 1) ÖNCE alt klasörler ---------------------------------------
        print("  · yonetim\\oneriler bırakma kutusuna çevriliyor")
        kod, cikti = icacls(o.oneriler, "/inheritance:r")
        if kod != 0:
            print(f"\n  icacls kullanılamıyor, test atlanıyor:\n  {cikti}")
            return 0
        kod, cikti = icacls(o.oneriler, "/grant", f"{KULLANICI}:{BIRAKMA_KUTUSU}")
        if kod != 0:
            print(f"\n  izin verilemedi, test atlanıyor:\n  {cikti}")
            return 0

        # Yönetim tarafının geri kalanı: personele yalnızca "izinleri gör".
        for hedef in (o.degerlendirme, gizli_not, o.yonetim_kitap):
            icacls(hedef, "/inheritance:r")
            icacls(hedef, "/grant", f"{KULLANICI}:(RC)")

        # --- 2) EN SON yonetim\ --------------------------------------------
        icacls(o.yonetim, "/inheritance:r")
        icacls(o.yonetim, "/grant", f"{KULLANICI}:{YONETIM_GECIS}")

        # --- kilitliyken neler engellenmiş? -------------------------------
        s.kontrol("yonetim\\ listelenemiyor",
                  _engellendi_mi(lambda: os.listdir(o.yonetim)))
        s.kontrol("Değerlendirme notu okunamıyor",
                  _engellendi_mi(lambda: open(gizli_not, encoding="utf-8").read()))
        s.kontrol("Değerlendirme notunun üzerine yazılamıyor",
                  _engellendi_mi(
                      lambda: open(gizli_not, "a", encoding="utf-8").write("x")))
        s.kontrol("Yönetim kitabı okunamıyor",
                  _engellendi_mi(lambda: open(o.yonetim_kitap, "rb").read()))
        s.kontrol("Öneriler klasörü listelenemiyor",
                  _engellendi_mi(lambda: os.listdir(o.oneriler)))

        # --- kilitliyken gerçek makroyu çalıştır --------------------------
        print("  · kilitli klasöre iki gönderim")
        numaralar = []
        with y.excel() as app:
            with y.kitap(app, o.oneri_kitap, salt_okunur=True) as wb:
                s.esit("Kök klasör kilitli yapıda da bulunuyor",
                       os.path.normcase(o.paylasim),
                       os.path.normcase(y.calistir(app, wb, "modAyar.KokKlasor")))
                for i in (1, 2):
                    numaralar.append(y.calistir(
                        app, wb, "modGonderim.TestGonderimi",
                        f"Kullanıcı {i}", f"1000{i}",
                        "Kilitli klasöre yazma denemesi.", f"Öneri {i}",
                        "Bir çözüm önerisi.", "Fayda"))

        s.kontrol("Kilitli klasöre gönderim yapılabildi",
                  all(n.startswith("PRJ-") for n in numaralar), str(numaralar))
        s.esit("İki gönderim iki farklı numara aldı", 2, len(set(numaralar)))

        if all(n.startswith("PRJ-") for n in numaralar):
            ilk = os.path.join(yil_klasoru, numaralar[0] + ".txt")
            s.kontrol("Bırakılan öneri okunamıyor",
                      _engellendi_mi(lambda: open(ilk, encoding="utf-8").read()))
            s.kontrol("Bırakılan öneri silinemiyor",
                      _engellendi_mi(lambda: os.remove(ilk)))
            s.kontrol("Yıl klasörü listelenemiyor",
                      _engellendi_mi(lambda: os.listdir(yil_klasoru)))

        # --- izinleri aç, sonuca bak --------------------------------------
        icacls(o.yonetim, "/reset", "/t", "/c", "/q")
        icacls(o.yonetim, "/grant", f"{KULLANICI}:(OI)(CI)(F)", "/t", "/c", "/q")

        yazilanlar = o.oneri_dosyalar()
        s.esit("Her iki öneri de gerçekten yazılmış", 2, len(yazilanlar))
        for yol in yazilanlar:
            kayit = y.kayit_oku(yol)
            s.kontrol(f"{os.path.basename(yol)} eksiksiz ve okunabilir",
                      kayit.get("kayit_sonu") == "1"
                      and kayit.get("oneri_no") in numaralar
                      and "Kilitli klasöre" in kayit.get("mevcut_durum", ""),
                      str(sorted(kayit)))
        s.kontrol("Değerlendirme notu bozulmadan duruyor",
                  "Gizli karar gerekçesi" in open(gizli_not, encoding="utf-8").read())

    finally:
        icacls(o.yonetim, "/reset", "/t", "/c", "/q")
        icacls(o.yonetim, "/grant", f"{KULLANICI}:(OI)(CI)(F)", "/t", "/c", "/q")
        shutil.rmtree(kok, ignore_errors=True)

    return s.bitir()


if __name__ == "__main__":
    sys.exit(calistir())
