# -*- coding: utf-8 -*-
r"""Eszamanlilik dogrulamasi -- sistemin en kritik testi.

Eski tasarimin iddiasi "kimse ortak bir dosyaya yazmaz, cakisma YAPISAL OLARAK
IMKANSIZDIR" idi. Veri artik TEK bir dosyanin icinde, dolayisiyla o iddia
gecerli degil; yerini su aldi:

    Herkes ayni dosyaya, KISA ve DISLAYICI bir islemle yazar. Kilit Excel'in
    kendi dosya kilididir. Catisan taraf geri cekilip yeniden dener.

Bu testin isi o cumleyi sinamaktir. Uc ayri risk olculur:

1) SERI YAZMA -- arka arkaya gonderim. Her gonderim dosyayi acip kaydedip
   kapatir; hicbir satirin kaybolmamasi ve sayacın atlamamasi beklenir.

2) GERCEK ESZAMANLILIK -- ayri sureclerde, ayri Excel ornekleri. Birden fazla
   personelin ayni anda gondermesi. Excel catismada HATA VERMEZ, dosyayi
   sessizce salt okunur acar; modDepo bunu catisma sayar. Bir satirin
   kaybolmasi ya da iki oneriye ayni numaranin verilmesi buradan cikardi.

3) EKIP KITABI ACIKKEN -- degerlendirme ekibi kitabi gun boyu acik tutar.
   Workbook_Open kilidi biraktigi icin personel yine de gonderebilmelidir;
   ayrica ekip ayni anda degerlendirme kaydedebilmelidir.

    python testler\test_eszamanlilik.py
"""

import multiprocessing
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yardimci as y

SERI_ADET = 20
SUREC_ADET = 3
SUREC_BASINA = 3
DEGERLENDIRME_ADET = 3


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
        kuyruk.put(("gonderim", etiket, numaralar, None))
    except Exception as hata:            # noqa: BLE001 -- alt surecten rapor
        kuyruk.put(("gonderim", etiket, [], repr(hata)))


def _surec_degerlendirme(kitap_yolu, numaralar, kuyruk):
    """Ekibi taklit eder: kitabi SALT OKUNUR acik tutar ve degerlendirme yazar.

    Salt okunur acmak sarttir. Uretimde bunu Workbook_Open yapar
    (modDepo.SaltOkunuraGec); testler EnableEvents=False ile kostugu icin o
    olay calismaz ve kitap acikca salt okunur acilmalidir. Aksi halde ekip gun
    boyu yazma kilidini tutar ve HICBIR personel oneri gonderemez.
    """
    try:
        sonuclar = []
        with y.excel() as app:
            with y.kitap(app, kitap_yolu, salt_okunur=True) as wb:
                for no in numaralar:
                    sonuclar.append(y.calistir(
                        app, wb, "modDegerlendirme.TestDegerlendirmesi",
                        no, "Değerlendirmede", f"{no} eşzamanlı olarak incelendi."))
        kuyruk.put(("degerlendirme", 0, sonuclar, None))
    except Exception as hata:            # noqa: BLE001
        kuyruk.put(("degerlendirme", 0, [], repr(hata)))


def _numaralar(oneriler):
    return [r["oneri_no"] for r in oneriler]


def _sirali_mi(numaralar, yil):
    """Numaralar 1'den baslayip boslugusuz mu artiyor?"""
    beklenen = [f"PRJ-{yil}-{i:04d}" for i in range(1, len(numaralar) + 1)]
    return sorted(numaralar) == beklenen


def calistir():
    s = y.Sonuc("Eşzamanlılık")
    baslangic_excel = y.excel_sayisi()

    with y.ortam() as o:
        # ------------------------------------------------------------------
        # 1. Seri yazma
        # ------------------------------------------------------------------
        print(f"  · arka arkaya {SERI_ADET} gönderim")
        numaralar = []
        with y.excel() as app:
            with y.kitap(app, o.oneri_kitap, salt_okunur=True) as wb:
                for i in range(SERI_ADET):
                    numaralar.append(y.calistir(
                        app, wb, "modGonderim.TestGonderimi",
                        "Seri Test", str(1000 + i),
                        "Arka arkaya gönderim denemesi.", f"Seri öneri {i}",
                        "Bir çözüm önerisi.", "Fayda"))

        hatali = [n for n in numaralar if not str(n).startswith("PRJ-")]
        s.kontrol("Seri gönderimlerin hepsi numara aldı", not hatali,
                  str(hatali[:3]))
        s.esit("Seri gönderimde numaralar benzersiz",
               SERI_ADET, len(set(numaralar)))
        s.kontrol("Seri gönderimde numaralar 0001'den boşluksuz gidiyor",
                  _sirali_mi(numaralar, y.yil()), str(sorted(numaralar)[:3]))

        oneriler, _ = y.depo_oku(o.yonetim_kitap)
        s.esit("Seri gönderimde hiçbir satır kaybolmadı",
               SERI_ADET, len(oneriler))
        s.esit("Depodaki numaralar döndürülenlerle aynı",
               sorted(numaralar), sorted(_numaralar(oneriler)))
        del app

        # ------------------------------------------------------------------
        # 2. Gercek eszamanlilik -- personel + ekip birlikte
        # ------------------------------------------------------------------
        print(f"  · {SUREC_ADET} personel süreci × {SUREC_BASINA} gönderim "
              f"+ 1 ekip süreci × {DEGERLENDIRME_ADET} değerlendirme (paralel)")
        onceki = len(oneriler)
        degerlendirilecek = sorted(numaralar)[:DEGERLENDIRME_ADET]

        kuyruk = multiprocessing.Queue()
        surecler = [
            multiprocessing.Process(target=_surec_gonderimi,
                                    args=(o.oneri_kitap, e, SUREC_BASINA, kuyruk))
            for e in range(1, SUREC_ADET + 1)
        ]
        surecler.append(multiprocessing.Process(
            target=_surec_degerlendirme,
            args=(o.yonetim_kitap, degerlendirilecek, kuyruk)))

        for p in surecler:
            p.start()

        sonuclar = [kuyruk.get(timeout=600) for _ in surecler]
        for p in surecler:
            p.join(timeout=120)

        hatalar = [f"{tur} {e}: {h}" for tur, e, _, h in sonuclar if h]
        s.kontrol("Hiçbir süreç hata vermedi", not hatalar, "; ".join(hatalar))

        paralel = [n for tur, _, liste, _ in sonuclar if tur == "gonderim"
                   for n in liste]
        beklenen = SUREC_ADET * SUREC_BASINA
        s.esit("Paralel gönderimlerin tamamı numara aldı",
               beklenen, len(paralel))
        s.kontrol("Paralel gönderimlerin hepsi başarılı",
                  all(str(n).startswith("PRJ-") for n in paralel),
                  str([n for n in paralel if not str(n).startswith("PRJ-")][:3]))
        s.esit("Paralel gönderimlerde numaralar benzersiz",
               beklenen, len(set(paralel)))

        deg_sonuc = [r for tur, _, liste, _ in sonuclar if tur == "degerlendirme"
                     for r in liste]
        s.kontrol("Ekip, personel yazarken değerlendirme kaydedebildi",
                  all(str(r).startswith("satir=") for r in deg_sonuc),
                  str(deg_sonuc))

        # --- Verilen numaralarla diskteki satirlar birebir ortusmeli -------
        #
        # Bu iki kontrol sessiz kaybi ADIYLA soyler. Bir yazicinin Save'i
        # basarisiz olur da hata yutulursa, cagirana numara BASARI gibi doner
        # ama satir diske hic inmez; ustelik sonraki yazici ayni "en buyuk + 1"
        # degerini hesaplayip AYNI numarayi verir. Sayilari karsilastirmak
        # bozuklugu gorur, bu satirlar hangi numaranin dustugunu gosterir.
        diskte = _numaralar(y.depo_oku(o.yonetim_kitap)[0])
        verilen = list(numaralar) + list(paralel)

        kayip = sorted(set(verilen) - set(diskte))
        s.kontrol("Numara verilen her gönderim gerçekten diske yazıldı",
                  not kayip, f"diskte olmayan numaralar: {kayip}")

        tekrarli = sorted({n for n in verilen if verilen.count(n) > 1})
        s.kontrol("Aynı numara iki kez verilmedi",
                  not tekrarli, f"tekrarlanan numaralar: {tekrarli}")

        # ------------------------------------------------------------------
        # 3. Depo tutarli mi?
        # ------------------------------------------------------------------
        oneriler, olaylar = y.depo_oku(o.yonetim_kitap)
        s.esit("Paralel gönderimde hiçbir satır kaybolmadı",
               onceki + beklenen, len(oneriler))
        s.esit("Değerlendirme olayları da kaybolmadı",
               DEGERLENDIRME_ADET, len(olaylar))

        depodakiler = _numaralar(oneriler)
        s.esit("Depoda tekrarlanan öneri numarası yok",
               len(depodakiler), len(set(depodakiler)))
        s.kontrol("Bütün numaralar 0001'den boşluksuz gidiyor",
                  _sirali_mi(depodakiler, y.yil()),
                  str(sorted(depodakiler)[-3:]))
        s.kontrol("Numaraların hepsi PRJ-YYYY-NNNN biçiminde",
                  all(re.fullmatch(r"PRJ-\d{4}-\d{4}", n) for n in depodakiler),
                  str([n for n in depodakiler
                       if not re.fullmatch(r"PRJ-\d{4}-\d{4}", n)][:3]))

        # Hicbir kayit bir digerinin uzerine yazilmamis olmali: her satirin
        # basligi kendine ozgu.
        basliklar = {r["oneri_basligi"] for r in oneriler}
        s.esit("Hiçbir satır bir diğerinin üzerine yazılmadı",
               len(oneriler), len(basliklar))
        s.kontrol("Her satırın zorunlu alanları dolu",
                  all(r["sema"] and r["tarih"] and r["ad_soyad"]
                      for r in oneriler),
                  str([r["oneri_no"] for r in oneriler
                       if not (r["sema"] and r["tarih"] and r["ad_soyad"])][:3]))

        y.okuyucuyu_kapat()
        s.kontrol("Geriye kilit dosyası kalmadı (~$ ve .kilit)",
                  not o.sahiplik_dosyalari(), str(o.sahiplik_dosyalari()))

    # Aranan sey "sifir Excel" degil, TESTIN sizdirmadigidir: baslangicta
    # baska bir isten kalan bir ornek kapanmak uzere olabilir.
    kalan = y.excel_sayisi_bekle(baslangic_excel)
    s.kontrol("Arkada görünmez Excel süreci kalmadı", kalan <= baslangic_excel,
              f"(başlangıç: {baslangic_excel}, bitiş: {kalan})")

    return s.bitir()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    sys.exit(calistir())
