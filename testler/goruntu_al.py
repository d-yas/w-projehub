# -*- coding: utf-8 -*-
r"""Ekranlarin PNG goruntusunu alir -- tasarim gozden gecirmesi icin.

Otomatik testler bir ekranin DOGRU calistigini gosterir, IYI GORUNDUGUNU
gosteremez. Bu arac her sayfayi oldugu gibi disa aktarir; boylece hizalama,
renk, kart ve dugme yerlesimi dosyayi elle acmadan gozden gecirilebilir.

    python testler\goruntu_al.py [hedef_klasor]
"""

import gc
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yardimci as y

XL_PDF = 0
DIKEY, YATAY = 1, 2

# (kitap, sayfa, aralik, yon) -- aralik ekranin tasarlanmis sinirlarini kapsar
GORUNUMLER = [
    ("oneri", "Giriş", "A1:G16", DIKEY),
    ("oneri", "Öneri Formu", "A1:G34", DIKEY),
    ("yonetim", "Giriş", "A1:G16", DIKEY),
    ("yonetim", "Liste", "A1:G22", YATAY),
    ("yonetim", "Değerlendirme", "A1:G48", DIKEY),
    ("yonetim", "Pano", "A1:M33", YATAY),
]


def _sayfa_disaver(wb, sayfa_adi, aralik, yon, hedef):
    """Bir ekrani PDF olarak disa aktarir.

    Ekran goruntusu (CopyPicture) pencerenin gercekten cizilmis olmasini
    ister ve gorunmez/otomasyon oturumlarinda basarisiz olur. PDF disa
    aktarma pencereye bagli degildir, sekilleri (dugmeleri) ve grafikleri
    icerir, ustelik vektoreldir.
    """
    ws = wb.Worksheets(sayfa_adi)
    eski = ws.Visible
    ws.Visible = -1
    try:
        ws.PageSetup.PrintArea = aralik
        ws.PageSetup.Orientation = yon
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.FitToPagesTall = False
        for kenar in ("LeftMargin", "RightMargin", "TopMargin", "BottomMargin"):
            setattr(ws.PageSetup, kenar, 14)          # ~0,5 cm
        ws.PageSetup.CenterHorizontally = True
        ws.ExportAsFixedFormat(XL_PDF, hedef, 0, True, False)
    finally:
        ws.Visible = eski
    return hedef


def calistir(hedef_klasor):
    os.makedirs(hedef_klasor, exist_ok=True)
    uretilen = []

    with y.ortam() as o:
        # Tek bir gorunur Excel oturumu: goruntu alma ekrandan gectigi icin
        # pencerenin acik olmasi gerekir.
        with y.excel(gorunur=True) as app:
            # Ornek veri: bos ekranlar tasarimi degerlendirmeye yetmez.
            with y.kitap(app, o.oneri_kitap, salt_okunur=True) as wb:
                ornekler = [
                    ("Ayşe Çağlar", "10045",
                     "Gişede müşteri sırası yoğun saatlerde 20 dakikayı buluyor.",
                     "Sıra yönetimi iyileştirmesi",
                     "Ön kontrol ekranı eklensin.", "İşlem başına 4 dakika"),
                    ("Mehmet Öz", "10046",
                     "Aynı evrak iki ayrı adımda taranıyor.", "Tek tarama kuralı",
                     "Tarama adımları birleştirilsin.", "Günde 1 saat"),
                    ("Zeynep Şahin", "10047",
                     "Onay için üç ayrı imza isteniyor.", "İmza sayısı azaltılsın",
                     "İki imza yeterli olsun.", "Dosya başına 1 gün"),
                    ("Can Yılmaz", "10048",
                     "Talepler tek kuyrukta birikiyor.", "Talep önceliklendirme",
                     "Kuyruk ikiye ayrılsın.", "Haftada 6 saat"),
                ]
                numaralar = [y.calistir(app, wb, "modGonderim.TestGonderimi", *e)
                             for e in ornekler]

            # IKISI DE SALT OKUNUR acilir. Yonetim kitabi burada yazma kipinde
            # tutulsaydi, degerlendirme yazan gizli Excel ornegi kilidi
            # bulamaz ve otuz saniye bekleyip pes ederdi.
            acik = {
                "oneri": y.ac(app, o.oneri_kitap),
                "yonetim": y.ac(app, o.yonetim_kitap),
            }
            wb = acik["yonetim"]

            y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                       numaralar[0], "Pilot Uygulamada",
                       "Şube müdürüyle görüşüldü.")
            y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                       numaralar[1], "Standartlaştırıldı",
                       "Tüm operasyona yaygınlaştırıldı.")
            y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                       numaralar[2], "Reddedildi",
                       "Mevzuat üç imzayı zorunlu kılıyor.")
            y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                       numaralar[3], "Değerlendirmede", "İnceleniyor.")

            # Kullanicinin bastigi dugmenin tam yolu: ozet satiri da dolsun.
            y.calistir(app, wb, "modKonsolide.OnerileriYenile")
            app.EnableEvents = False
            y.calistir(app, wb, "modDegerlendirme.OneriyiAc", numaralar[0])

            for kitap_anahtar, sayfa, aralik, yon in GORUNUMLER:
                ad = f"{kitap_anahtar}-{sayfa.replace(' ', '-')}.pdf"
                hedef = os.path.join(hedef_klasor, ad)
                _sayfa_disaver(acik[kitap_anahtar], sayfa, aralik, yon, hedef)
                uretilen.append(hedef)
                print(f"  ✓ {ad}")

            for k in acik.values():
                k.Saved = True
                k.Close(SaveChanges=False)
            acik.clear()
            y.okuyucuyu_kapat()

        # Excel sureci, kendisine ait son COM vekili birakilana kadar olmez;
        # "with ... as app" blogu bitse de degisken bagli kalir.
        del app, wb
        gc.collect()

    return uretilen


if __name__ == "__main__":
    klasor = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "goruntuler")
    print("Ekran görüntüleri alınıyor…")
    for y_ in calistir(klasor):
        pass
    print(f"\nKlasör: {klasor}")
