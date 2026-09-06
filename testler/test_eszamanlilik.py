# -*- coding: utf-8 -*-
r"""Eszamanlilik dogrulamasi.

Sistemin temel iddiasi: "Kimse ortak bir dosyaya yazmaz, bu yuzden cakisma
onlenmis degil YAPISAL OLARAK IMKANSIZDIR." Bu test o iddiayi sinar.

Iki ayri risk olculur:

1) SERI PATLAMA -- ayni saniye icinde arka arkaya gonderim.
   Basvuru numarasi zaman damgasi + rastgele ek oldugu icin, en gercekci
   catisma riski "ayni saniye" penceresidir. Tek surecte hizli bir dizi
   gonderim bu pencereyi doldurur.

2) GERCEK EScAMANLILIK -- ayri sureclerde, ayri Excel ornekleri.
   Birden fazla kullanicinin ayni anda gondermesini taklit eder. Hicbir
   kaydin kaybolmamasi ve hicbir dosya adinin cakismamasi beklenir.

    python testler\test_eszamanlilik.py
"""

import multiprocessing
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yardimci as y

SERI_ADET = 30
SUREC_ADET = 3
SUREC_BASINA = 3


def _surec_gonderimi(kitap_yolu, etiket, adet, kuyruk):
    """Ayri bir surecte kendi Excel ornegiyle gonderim yapar."""
    try:
        numaralar = []
        with y.excel() as app:
            # Uretimde bu dosya tum personel icin salt okunurdur; birden fazla
            # kullanicinin ayni anda acabilmesinin sarti budur.
            with y.kitap(app, kitap_yolu, salt_okunur=True) as wb:
                for i in range(adet):
                    numaralar.append(y.calistir(
                        app, wb, "modGonderim.TestGonderimi",
                        f"Kullanıcı {etiket}", f"90{etiket}{i}",
                        f"{etiket} numaralı kullanıcının sorunu.",
                        f"{etiket}-{i} önerisi", "Bir çözüm önerisi.", "Fayda"))
        kuyruk.put((etiket, numaralar, None))
    except Exception as hata:            # noqa: BLE001 -- alt surecten rapor
        kuyruk.put((etiket, [], repr(hata)))


def calistir():
    s = y.Sonuc("Eşzamanlılık")

    with y.ortam(yonetim_de=False) as o:
        # ------------------------------------------------------------------
        # 1. Seri patlama
        # ------------------------------------------------------------------
        print(f"  · aynı saniye penceresinde {SERI_ADET} gönderim")
        numaralar = []
        with y.excel() as app:
            with y.kitap(app, o.oneri_kitap) as wb:
                for i in range(SERI_ADET):
                    numaralar.append(y.calistir(
                        app, wb, "modGonderim.TestGonderimi",
                        "Seri Test", str(1000 + i),
                        "Arka arkaya gönderim denemesi.", f"Seri öneri {i}",
                        "Bir çözüm önerisi.", "Fayda"))

        s.esit("Seri gönderimde numaralar benzersiz",
               SERI_ADET, len(set(numaralar)))
        s.esit("Seri gönderimde hiçbir kayıt kaybolmadı",
               SERI_ADET, len(o.oneri_dosyalar()))

        # ------------------------------------------------------------------
        # 2. Gercek eszamanlilik -- ayri surecler, ayri Excel ornekleri
        # ------------------------------------------------------------------
        print(f"  · {SUREC_ADET} ayrı süreç × {SUREC_BASINA} gönderim (paralel)")
        onceki = len(o.oneri_dosyalar())

        kuyruk = multiprocessing.Queue()
        surecler = [
            multiprocessing.Process(target=_surec_gonderimi,
                                    args=(o.oneri_kitap, e, SUREC_BASINA, kuyruk))
            for e in range(1, SUREC_ADET + 1)
        ]
        for p in surecler:
            p.start()

        sonuclar = [kuyruk.get(timeout=300) for _ in surecler]
        for p in surecler:
            p.join(timeout=60)

        hatalar = [f"süreç {e}: {h}" for e, _, h in sonuclar if h]
        s.kontrol("Hiçbir süreç hata vermedi", not hatalar, "; ".join(hatalar))

        paralel_numaralar = [n for _, liste, _ in sonuclar for n in liste]
        beklenen = SUREC_ADET * SUREC_BASINA
        s.esit("Paralel gönderimlerin tamamı numara aldı",
               beklenen, len(paralel_numaralar))
        s.esit("Paralel gönderimlerde numaralar benzersiz",
               beklenen, len(set(paralel_numaralar)))

        dosyalar = o.oneri_dosyalar()
        s.esit("Paralel gönderimde hiçbir kayıt kaybolmadı",
               onceki + beklenen, len(dosyalar))

        adlar = [os.path.basename(d) for d in dosyalar]
        s.esit("Dosya adları çakışmadı", len(adlar), len(set(adlar)))

        # ------------------------------------------------------------------
        # 3. Her kayit eksiksiz yazildi
        # ------------------------------------------------------------------
        eksik = [os.path.basename(d) for d in dosyalar
                 if "kayit_sonu" not in y.kayit_oku(d)]
        s.kontrol("Bütün kayıtlar sonu işaretli (yarım yazılan yok)",
                  not eksik, str(eksik))
        yabanci = []
        for dizin, _, dosya_listesi in os.walk(o.oneriler):
            yabanci += [d for d in dosya_listesi if not d.endswith(".txt")]
        s.kontrol("Geride geçici dosya kalmadı", not yabanci, str(yabanci))

        # Her kayit okunabilir ve kendi numarasini tasiyor mu?
        bozuk = []
        for d in dosyalar:
            kayit = y.kayit_oku(d)
            beklenen_no = os.path.splitext(os.path.basename(d))[0]
            if kayit.get("oneri_no") != beklenen_no:
                bozuk.append(os.path.basename(d))
        s.kontrol("Her dosya kendi öneri numarasını taşıyor", not bozuk, str(bozuk))

        # Kisa numara, kayit kaybi riski getirmemeli: butun gonderimlerin
        # icerigi ayri ayri okunabilmeli ve hepsi farkli olmali.
        basliklar = {y.kayit_oku(d).get("oneri_basligi", "") for d in dosyalar}
        s.esit("Hiçbir kayıt bir diğerinin üzerine yazılmadı",
               len(dosyalar), len(basliklar))

    return s.bitir()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(calistir())
