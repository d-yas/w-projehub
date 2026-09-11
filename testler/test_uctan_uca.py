# -*- coding: utf-8 -*-
r"""Uctan uca dogrulama -- gercek Excel, gercek makrolar, gercek dosya.

Bu bir taklit (mock) testi degildir: uretilen .xlsm dosyalarini Excel'de acar,
VBA makrolarini calistirir ve sonuclari DISKTEKI yonetim kitabindan okur.

TESTIN GOREMEDIGI (bilincli sinir):
Makrolar COM uzerinden cagrilir. Bu yol Excel'in makro guvenlik ayarini,
"Icerigi Etkinlestir" uyarisini ve dugmelerin gorsel yerlesimini HIC gormez.
Yani burasi "kod dogru mu" sorusunu yanitlar, "kullanici bu dosyayi actiginda
ne olur" sorusunu yanitlamaz. Kurulumdan sonra TALIMATNAME.md madde 6'daki elle
kontrol listesi mutlaka bir kez uygulanmalidir.

    python testler\test_uctan_uca.py
"""

import os
import re
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yardimci as y


def calistir():
    s = y.Sonuc("Uçtan uca")
    baslangic_excel = y.excel_sayisi()

    with y.ortam() as o:
        with y.excel() as app:
            # Sira onemli: sayi bekleyen testler once calisir, veri ekleyen
            # ekran testleri en sona birakilir.
            _gonderim(s, app, o)
            _konsolidasyon(s, app, o)
            _degerlendirme(s, app, o)
            _gostergeler(s, app, o)
            _degerlendirme_ekrani(s, app, o)
            _takip(s, app, o)
            _form_ekrani(s, app, o)
            _kitap_kaydedilmez(s, app, o)
            _yedek(s, app, o)
            _kilitliyken_gonderim(s, app, o)

            # Okuyucu ornegi COM apartmani hala acikken kapatilmali.
            y.okuyucuyu_kapat()

        # "with ... as app" blogu bitince degisken BAGLI KALIR ve Excel
        # sureci, kendisine ait son COM vekili birakilana kadar olmez.
        del app

        s.kontrol("Geriye kilit dosyası kalmadı (~$ ve .kilit)",
                  not o.sahiplik_dosyalari(), str(o.sahiplik_dosyalari()))

    # modDepo her yazma icin gizli bir Excel ornegi acar. Kapatilmayan bir
    # ornek dosya kilidini tutmaya devam eder ve bir sonraki gonderimi otuz
    # saniye bekletir; sayilarin esit olmasi bunu disarida birakir.
    # Aranan sey "sifir Excel" degil, TESTIN sizdirmadigidir: baslangicta
    # baska bir isten kalan bir ornek kapanmak uzere olabilir.
    kalan = y.excel_sayisi_bekle(baslangic_excel)
    s.kontrol("Arkada görünmez Excel süreci kalmadı", kalan <= baslangic_excel,
              f"(başlangıç: {baslangic_excel}, bitiş: {kalan})")

    return s.bitir()


# ==========================================================================
#  1. Gonderim
# ==========================================================================
def _gonderim(s, app, o):
    print("  · gönderim akışı")
    with y.kitap(app, o.oneri_kitap, salt_okunur=True) as wb:
        # ThisWorkbook yalnizca bir olay tetiklendiginde derlenir; oradaki bir
        # hata aksi halde ilk kullanicida ortaya cikar. Bu cagri derlemeye zorlar.
        s.kontrol("Öneri kitabının ThisWorkbook modülü derleniyor",
                  y.derleme_sinamasi(app, wb))

        s.esit("Türkçe karakterler VBA kaynağına doğru aktarıldı",
               "Ölçümleniyor|Standartlaştırıldı|Hızlı Kazanım|İşi|ĞÜŞİÖÇ|ğüşıöç",
               y.calistir(app, wb, "modTasarim.KodlamaSinamasi"))

        s.esit("Kök klasör kitabın konumundan bulundu",
               os.path.normcase(o.paylasim),
               os.path.normcase(y.calistir(app, wb, "modAyar.KokKlasor")))

        s.esit("Veri deposunun yolu yönetim kitabıdır",
               os.path.normcase(o.yonetim_kitap),
               os.path.normcase(y.calistir(app, wb, "modAyar.YonetimKitapYolu")))

        # Form sayfasinda birim/israf alanlari kalmadi.
        ws = wb.Worksheets("Öneri Formu")
        for kalkan in ("frm_birim", "frm_israf"):
            var_mi = True
            try:
                ws.Range(kalkan)
            except Exception:
                var_mi = False
            s.kontrol(f"Kaldırılan alan formda yok: {kalkan}", not var_mi)

        o.no1 = y.calistir(
            app, wb, "modGonderim.TestGonderimi",
            "Ayşe Çağlar", "10045",
            "Müşteri gişede sıra bekliyor.\nİkinci satır: yoğun saatlerde 20 dk.",
            "Sıra yönetimi iyileştirmesi",
            "Ön kontrol ekranı eklensin.",
            "İşlem başına 4 dakika")
        o.no2 = y.calistir(
            app, wb, "modGonderim.TestGonderimi",
            "Mehmet Öz", "10046",
            "Aynı evrak iki kez taranıyor.", "Tek tarama kuralı",
            "Tarama adımı birleştirilsin.", "Günde 1 saat")
        o.no3 = y.calistir(
            app, wb, "modGonderim.TestGonderimi",
            "Zeynep Şahin", "10047",
            "Onay için üç imza isteniyor.", "İmza sayısı azaltılsın",
            "İki imza yeterli olsun.", "Dosya başına 1 gün")

    oneriler, olaylar = y.depo_oku(o.yonetim_kitap)
    s.esit("Üç gönderim depoya üç satır yazdı", 3, len(oneriler))
    s.esit("Henüz değerlendirme olayı yok", 0, len(olaylar))

    s.kontrol("Öneri numaraları benzersiz", len({o.no1, o.no2, o.no3}) == 3)
    s.kontrol("Öneri numarası PRJ-YYYY-NNNN biçiminde",
              re.fullmatch(r"PRJ-\d{4}-\d{4}", o.no1) is not None, o.no1)
    s.esit("Numaralar sırayla verildi", [f"PRJ-{y.yil()}-000{i}" for i in (1, 2, 3)],
           [o.no1, o.no2, o.no3])

    s.esit("Depodaki satır sırası gönderim sırasıyla aynı",
           [o.no1, o.no2, o.no3], [r["oneri_no"] for r in oneriler])

    kayit = oneriler[0]
    s.esit("Türkçe karakterler kayıtta bozulmadı", "Ayşe Çağlar", kayit["ad_soyad"])
    s.esit("Şema sürümü yazıldı", "4", kayit["sema"])
    s.kontrol("Çok satırlı alan hücrede satır sonu olarak duruyor",
              kayit["mevcut_durum"].count("\n") == 1, repr(kayit["mevcut_durum"]))
    s.esit("Sicil numarası metin olarak korundu (baştaki sıfırlar yenmez)",
           "10045", kayit["sicil_no"])
    s.kontrol("Gönderim kaydında durum alanı YOK (durum yalnızca olaylardan gelir)",
              "durum" not in kayit, sorted(kayit))
    s.kontrol("Kayıtta birim ve israf alanları yok",
              "birim" not in kayit and "israf_turu" not in kayit, sorted(kayit))
    s.kontrol("İzlenebilirlik alanları yazıldı",
              len(kayit.get("gonderen_kullanici", "")) > 0
              and len(kayit.get("gonderen_bilgisayar", "")) > 0)
    s.kontrol("Tarih ISO biçiminde",
              len(kayit["tarih"]) == 19 and kayit["tarih"][10] == "T", kayit["tarih"])


# ==========================================================================
#  2. Konsolidasyon
# ==========================================================================
def _konsolidasyon(s, app, o):
    print("  · konsolidasyon")

    with y.kitap(app, o.yonetim_kitap, salt_okunur=True) as wb:
        s.kontrol("Yönetim kitabının ThisWorkbook modülü derleniyor",
                  y.derleme_sinamasi(app, wb))

        adet = y.calistir(app, wb, "modKonsolide.VeriyiKur")
        s.esit("Üç öneri okundu", 3, adet)

        s.esit("Yeniden okuma aynı sonucu veriyor", 3,
               y.calistir(app, wb, "modKonsolide.VeriyiKur"))

        # Dugmenin cagirdigi tam yol: gizli bir Excel ornegi acar, depoyu ceker,
        # tabloyu ve panoyu kurar. Daha once burada sessizce kilitleniyordu.
        y.calistir(app, wb, "modKonsolide.OnerileriYenile")
        app.EnableEvents = False        # HizliModKapa olaylari geri aciyor
        s.esit("“Önerileri Yenile” düğmesi hatasız çalıştı", "",
               y.son_mesaj(app, wb))
        s.kontrol("Liste özeti yazıldı",
                  "öneri okundu" in str(
                      wb.Worksheets("Liste").Range("liste_ozet").Value or ""),
                  str(wb.Worksheets("Liste").Range("liste_ozet").Value))

        ws = wb.Worksheets("Liste")
        s.esit("Liste tablosuna üç satır yazıldı", 3,
               sum(1 for r in range(9, 20)
                   if str(ws.Cells(r, 2).Value or "").startswith("PRJ-")))

        s.esit("Yeni öneriler 'Yeni' durumuyla başlıyor", "Yeni",
               str(ws.Cells(9, 6).Value or ""))

        # Sahipsiz bir olay -- oneri numarasi hicbir gonderime denk gelmiyor.
        # Konsolidasyon onu sessizce yok saymali, patlamamali.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   "PRJ-1900-9999", "Planlandı", "Sahipsiz olay.")
        y.calistir(app, wb, "modKonsolide.OnerileriYenile")
        app.EnableEvents = False
        s.esit("Sahipsiz olay öneri sayısını değiştirmedi", 3,
               y.calistir(app, wb, "modKonsolide.VeriyiKur"))
        s.esit("Sahipsiz olaydan sonra da hata yok", "", y.son_mesaj(app, wb))

    oneriler, olaylar = y.depo_oku(o.yonetim_kitap)
    s.esit("Sahipsiz olay yine de depoda duruyor (silinmez)", 1, len(olaylar))
    s.esit("Gönderim satırları olduğu gibi duruyor", 3, len(oneriler))


# ==========================================================================
#  3. Degerlendirme (ekle-only)
# ==========================================================================
def _degerlendirme(s, app, o):
    print("  · değerlendirme ve geçmiş")

    with y.kitap(app, o.yonetim_kitap, salt_okunur=True) as wb:
        onceki = len(y.depo_oku(o.yonetim_kitap)[1])

        # Birinci degerlendirme: planlanir.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no1, "Planlandı",
                   "Şube müdürüyle görüşüldü, uygulanabilir.")

        _, olaylar = y.depo_oku(o.yonetim_kitap)
        s.esit("İlk değerlendirme bir olay satırı ekledi", onceki + 1, len(olaylar))
        ilk_satir = dict(olaylar[-1])

        y.calistir(app, wb, "modKonsolide.OnerileriYenile")
        app.EnableEvents = False
        s.esit("Konsolidasyon öneri sayısını değiştirmedi", 3,
               y.calistir(app, wb, "modKonsolide.VeriyiKur"))

        veri = _veri_satiri(wb, o.no1)
        s.esit("Durum olaydan türetildi", "Planlandı", veri["durum"])
        s.esit("Karar notu uygulandı", "Şube müdürüyle görüşüldü, uygulanabilir.",
               veri["karar_notu"])

        # Ikinci degerlendirme: YALNIZCA durum degisir, not verilmez.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no1, "Pilot Uygulamada", "")

        _, olaylar = y.depo_oku(o.yonetim_kitap)
        s.esit("İkinci değerlendirme ikinci satırı ekledi", onceki + 2, len(olaylar))
        s.esit("İlk olay satırı olduğu gibi duruyor (ekle-only)",
               ilk_satir, dict(olaylar[-2]))

        y.calistir(app, wb, "modKonsolide.OnerileriYenile")
        app.EnableEvents = False
        veri = _veri_satiri(wb, o.no1)
        s.esit("Son olay durumu belirledi", "Pilot Uygulamada", veri["durum"])
        s.esit("Kısmi olay önceki karar notunu SİLMEDİ",
               "Şube müdürüyle görüşüldü, uygulanabilir.", veri["karar_notu"])
        s.esit("Olay sayısı iki", 2, veri["olay_sayisi"])

        # Diger iki oneri: biri reddedilir, biri planlanir ama uygulanmaz.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no2, "Planlandı", "Kabul edildi, BT planına alındı.")
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no3, "Reddedildi", "Mevzuat üç imzayı zorunlu kılıyor.")

        y.calistir(app, wb, "modKonsolide.OnerileriYenile")
        app.EnableEvents = False
        s.esit("İkinci öneri planlandı olarak işlendi",
               "Planlandı", _veri_satiri(wb, o.no2)["durum"])
        s.esit("Üçüncü öneri reddedildi olarak işlendi",
               "Reddedildi", _veri_satiri(wb, o.no3)["durum"])

        s.esit("Bir önerinin geçmişi iki olaydan oluşuyor", 2,
               y.calistir(app, wb, "modDegerlendirme.TestOlaySayisi", o.no1))


def _veri_satiri(wb, oneri_no):
    """Gizli Veri sayfasindan bir satiri sozluk olarak okur."""
    ws = wb.Worksheets("Veri")
    adet = int(ws.Range("veri_adet").Value or 0)
    for r in range(2, adet + 2):
        if str(ws.Cells(r, 1).Value or "") == oneri_no:
            return {
                "durum": str(ws.Cells(r, 9).Value or ""),
                "karar_notu": str(ws.Cells(r, 13).Value or ""),
                "olay_sayisi": int(ws.Cells(r, 14).Value or 0),
            }
    raise AssertionError(f"Veri sayfasında bulunamadı: {oneri_no}")


# ==========================================================================
#  4. Gostergeler
# ==========================================================================
def _gostergeler(s, app, o):
    print("  · göstergeler")

    with y.kitap(app, o.yonetim_kitap, salt_okunur=True) as wb:
        y.calistir(app, wb, "modKonsolide.OnerileriYenile")
        app.EnableEvents = False
        ham = y.calistir(app, wb, "modPano.TestGostergeleri")
        g = dict(p.split("=", 1) for p in ham.split(";"))

        s.esit("Toplam öneri", "3", g["toplam"])

        # Panodaki kartlarin govdesi hucre degil SEKILDIR; rakam hucrede
        # durur ama ekranda sekil gorunur. Ikisi ayrisirsa kullanici yanlis
        # sayiya bakar -- bu kontrol o ayrisan durumu yakalar.
        s.esit("Kart şekli toplam öneriyi gösteriyor", "3",
               y.calistir(app, wb, "modPano.TestKartMetni", "pano_toplam"))
        s.esit("Bekleyen öneri yok (üçü de sonuçlandı)", "0", g["bekleyen"])
        s.esit("Uygulamaya geçmiş bir öneri", "1", g["uygulanan"])
        s.esit("Bu ay gelen öneri sayısı", "3", g["bu_ay"])

        # "Planlandi" henuz uygulanmis sayilmaz; ikinci oneri uygulamaya
        # alininca sayac artmalidir.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no2, "Standartlaştırıldı", "Tüm operasyona yaygınlaştırıldı.")
        y.calistir(app, wb, "modKonsolide.OnerileriYenile")
        app.EnableEvents = False
        ham = y.calistir(app, wb, "modPano.TestGostergeleri")
        g = dict(p.split("=", 1) for p in ham.split(";"))

        s.esit("Uygulamaya geçince sayaç arttı", "2", g["uygulanan"])

        # Grafik kaynak verisi
        pv = wb.Worksheets("PanoVeri")
        s.esit("Grafik verisi: durum dağılımı ilk satır etiketi",
               "Yeni", str(pv.Cells(2, 1).Value or ""))
        toplam_durum = sum(int(pv.Cells(r, 2).Value or 0) for r in range(2, 10))
        s.esit("Grafik verisi: durum dağılımı toplamı öneri sayısına eşit",
               3, toplam_durum)

        # Grafiklerin HİÇBİR kategoriyi düşürmediği doğrulanır. Başlıksız bir
        # kaynak aralığı verilirse Excel ilk veri satırını başlık sanar ve o
        # kategori sessizce grafikten çıkar.
        pano = wb.Worksheets("Pano")
        nokta_sayilari = sorted(
            co.Chart.SeriesCollection(1).Points().Count for co in pano.ChartObjects())
        s.esit("Grafiklerde kategori sayıları eksiksiz (durum 8, ay 12)",
               [8, 12], nokta_sayilari)
        s.kontrol("Grafiklerde başlık ve gösterge kapalı",
                  all(not co.Chart.HasTitle and not co.Chart.HasLegend
                      for co in pano.ChartObjects()))


# ==========================================================================
#  5. Degerlendirme ekraninin GERCEK dugme yolu
# ==========================================================================
def _degerlendirme_ekrani(s, app, o):
    print("  · değerlendirme ekranı (Yükle / Kaydet düğmeleri)")

    with y.kitap(app, o.yonetim_kitap, salt_okunur=True) as wb:
        y.calistir(app, wb, "modKonsolide.OnerileriYenile")
        app.EnableEvents = False
        ws = wb.Worksheets("Değerlendirme")

        y.calistir(app, wb, "modDegerlendirme.OneriyiAc", o.no1)
        s.esit("Öneri ekrana hatasız yüklendi", "", y.son_mesaj(app, wb))
        s.esit("Öneri numarası ekrana geldi", o.no1,
               str(ws.Range("dg_oneri_no").Value or ""))
        s.kontrol("Öneri metni ekrana geldi",
                  "Sıra yönetimi" in str(ws.Range("dg_baslik").Value or ""),
                  str(ws.Range("dg_baslik").Value))
        s.esit("Geçmiş listesinde en yeni değerlendirme üstte",
               "Pilot Uygulamada", str(ws.Cells(34, 4).Value or ""))

        # Reddetme gerekcesiz kabul edilmemeli.
        _ekrana_yaz(ws, {"dg_yeni_durum": "Reddedildi", "dg_not": ""})
        y.calistir(app, wb, "modDegerlendirme.DegerlendirmeKaydet")
        mesaj = y.son_mesaj(app, wb)
        s.kontrol("Gerekçesiz reddetme engellendi",
                  mesaj.startswith("hata:") and "gerekçe" in mesaj, mesaj)

        # Gecerli bir kayit
        onceki = len(y.depo_oku(o.yonetim_kitap)[1])
        _ekrana_yaz(ws, {
            "dg_yeni_durum": "Ölçümleniyor",
            "dg_not": "Pilot sonuçları ölçülmeye başlandı."})
        y.calistir(app, wb, "modDegerlendirme.DegerlendirmeKaydet")
        app.EnableEvents = False        # Kaydet sonrasi yenileme olaylari geri acar

        s.esit("Kaydet düğmesi yeni olay satırı ekledi",
               onceki + 1, len(y.depo_oku(o.yonetim_kitap)[1]))
        s.kontrol("Kayıt sonrası onay bandı yazıldı",
                  "kaydedildi" in str(ws.Range("dg_bant").Value or ""),
                  str(ws.Range("dg_bant").Value))

        # Kaydet, ekrani KENDISI tazelemeli: ayrica Yenile'ye basmak gerekmez.
        veri = _veri_satiri(wb, o.no1)
        s.esit("Ekrandan kaydedilen durum hemen uygulandı", "Ölçümleniyor",
               veri["durum"])
        s.esit("Ekrandan kaydedilen karar notu uygulandı",
               "Pilot sonuçları ölçülmeye başlandı.", veri["karar_notu"])


def _ekrana_yaz(ws, degerler):
    """Korumalı sayfadaki giriş alanlarını doldurur."""
    y.korumasiz(ws)
    for ad, deger in degerler.items():
        ws.Range(ad).Value = deger


# ==========================================================================
#  5b. Takip ekrani -- personelin kendi onerisini sorgulamasi
#
#  Takip kitabi PAROLASIZDIR: depo oraya kopyalanmaz, yalnizca sorgulanan
#  onerinin satirlari gelir. Numaralar sirali ve tahmin edilebilir oldugu
#  icin sicil de tutmalidir; ekibin karar notlari oneri sahibine gosterilmez.
#
#  Bu noktada birinci oneri Planlandı -> Pilot Uygulamada -> Ölçümleniyor
#  gecmisine, ucuncusu gerekceli bir Reddedildi'ye sahiptir.
# ==========================================================================
def _takip(s, app, o):
    print("  · takip ekranı (Sorgula / Temizle düğmeleri)")

    with y.kitap(app, o.takip_kitap, salt_okunur=True) as wb:
        s.kontrol("Takip kitabının ThisWorkbook modülü derleniyor",
                  y.derleme_sinamasi(app, wb))
        s.esit("Takip kitabı veri deposunu buluyor",
               os.path.normcase(o.yonetim_kitap),
               os.path.normcase(y.calistir(app, wb, "modAyar.YonetimKitapYolu")))
        s.kontrol("Takip kitabında depo sayfası yok",
                  not ({"Oneriler", "Olaylar", "Veri"}
                       & {w.Name for w in wb.Worksheets}))

        ws = wb.Worksheets("Takip")

        # Numara kucuk harfle ve bosluklu yazilir; kullanicilar boyle yazar.
        _sorgula(app, wb, ws, f"  {o.no1.lower()} ", "10045")
        s.esit("Numara + sicil ile sorgu hatasız", "", y.son_mesaj(app, wb))
        s.esit("Küçük harfle yazılan numara bulundu", o.no1,
               _hucre(ws, "tk_sonuc_no"))
        s.esit("Güncel durum listedekiyle aynı", "Ölçümleniyor",
               _hucre(ws, "tk_durum"))
        s.kontrol("Öneri başlığı gösterildi",
                  "Sıra yönetimi" in _hucre(ws, "tk_baslik"), _hucre(ws, "tk_baslik"))
        s.kontrol("Gönderim tarihi okunur biçimde (gg.aa.yyyy)",
                  re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", _hucre(ws, "tk_tarih"))
                  is not None, _hucre(ws, "tk_tarih"))
        s.kontrol("Onay bandı yazıldı ve işareti bozulmadı (✓, '?' değil)",
                  _hucre(ws, "tk_bant").startswith("✓")
                  and "bulundu" in _hucre(ws, "tk_bant"), _hucre(ws, "tk_bant"))

        gecmis = [str(ws.Cells(r, 3).Value or "") for r in range(31, 41)]
        s.esit("Durum geçmişi değişiklikleri en yeniden eskiye gösteriyor",
               ["Ölçümleniyor", "Pilot Uygulamada", "Planlandı", "Yeni"],
               [g for g in gecmis if g])
        s.kontrol("Geçmişin her satırında durumun anlamı var",
                  all(str(ws.Cells(r, 4).Value or "") for r in range(31, 35)))
        metin = _sayfa_metni(ws)
        s.kontrol("Karar notları takip ekranında görünmüyor",
                  not any(n in metin for n in ("Şube müdürüyle", "Pilot sonuçları")))

        # Yanlis sicil ile olmayan numara AYNI cevabi almali; aksi halde ekran
        # hangi numaralarin var oldugunu ele verir.
        _sorgula(app, wb, ws, o.no1, "99999")
        yanlis_sicil = y.son_mesaj(app, wb)
        s.kontrol("Yanlış sicil ile sonuç gösterilmedi",
                  yanlis_sicil.startswith("hata:") and not _hucre(ws, "tk_durum"),
                  yanlis_sicil)
        s.esit("Önceki sorgunun sonucu ekrandan silindi", "",
               _hucre(ws, "tk_sonuc_no"))
        s.esit("Önceki sorgunun geçmişi ekrandan silindi", "",
               str(ws.Cells(31, 3).Value or ""))

        _sorgula(app, wb, ws, f"PRJ-{y.yil()}-9999", "10045")
        s.esit("Olmayan numara ile yanlış sicil aynı mesajı alıyor",
               yanlis_sicil, y.son_mesaj(app, wb))

        # Gerekce depoda var ama oneri sahibine gosterilmez.
        _sorgula(app, wb, ws, o.no3, "10047")
        s.esit("Reddedilen öneri sorgulandı", "Reddedildi", _hucre(ws, "tk_durum"))
        s.kontrol("Ret gerekçesi takip ekranında görünmüyor",
                  "Mevzuat" not in _sayfa_metni(ws))

        _sorgula(app, wb, ws, o.no2, "")
        mesaj = y.son_mesaj(app, wb)
        s.kontrol("Sicil boşken sorgu yapılmadı",
                  mesaj.startswith("hata:") and "sicil" in mesaj, mesaj)

        _sorgula(app, wb, ws, o.no2, "10046")
        y.calistir(app, wb, "modTakip.Temizle")
        s.esit("Temizle düğmesi hata vermedi", "", y.son_mesaj(app, wb))
        dolu = [ad for ad in ("tk_no", "tk_sicil", "tk_sonuc_no", "tk_durum",
                              "tk_bant") if _hucre(ws, ad)]
        s.kontrol("Temizle girişleri ve sonucu boşalttı", not dolu, str(dolu))


def _sorgula(app, wb, ws, no, sicil):
    _ekrana_yaz(ws, {"tk_no": no, "tk_sicil": sicil})
    y.calistir(app, wb, "modTakip.Sorgula")
    app.EnableEvents = False        # HizliModKapa olaylari geri aciyor


def _hucre(ws, ad):
    return str(ws.Range(ad).Value or "")


def _sayfa_metni(ws):
    """Sayfada yazan her seyi tek metin olarak dondurur."""
    return "\n".join(str(h) for satir in (ws.UsedRange.Value or ())
                     for h in satir if h is not None)


# ==========================================================================
#  6. Form ekraninin GERCEK dugme yolu
# ==========================================================================
def _form_ekrani(s, app, o):
    print("  · form ekranı (Gönder / Temizle düğmeleri)")

    alanlar = {
        "frm_ad_soyad": "Elif Karaduman",
        "frm_sicil_no": "10099",
        "frm_mevcut": "Rapor almak için üç farklı ekran açmak gerekiyor.",
        "frm_baslik": "Tek ekranda rapor",
        "frm_cozum": "Üç ekran tek bir görünümde birleştirilsin.",
        "frm_fayda": "Rapor başına 10 dakika",
    }

    onceki = len(y.depo_oku(o.yonetim_kitap)[0])
    with y.kitap(app, o.oneri_kitap, salt_okunur=True) as wb:
        ws = wb.Worksheets("Öneri Formu")
        ws.Visible = -1
        _ekrana_yaz(ws, alanlar)

        y.calistir(app, wb, "modGonderim.OneriGonder")
        mesaj = y.son_mesaj(app, wb)

        s.kontrol("Gönder düğmesi başarı mesajı verdi",
                  mesaj.startswith("bilgi:") and "Öneri numaranız" in mesaj, mesaj)
        s.kontrol("Onay bandı öneri numarasını gösteriyor, işareti bozulmadı",
                  "PRJ-" in str(ws.Range("frm_bant").Value or "")
                  and str(ws.Range("frm_bant").Value or "").startswith("✓"),
                  str(ws.Range("frm_bant").Value))

        # Gonderim sonrasi form kendiliginden temizlenmis olmali.
        dolu = [ad for ad in alanlar
                if str(ws.Range(ad).Value or "").strip()]
        s.kontrol("Gönderimden sonra form temizlendi", not dolu, str(dolu))

        # "Temizle" dugmesi de tek basina calismali.
        _ekrana_yaz(ws, alanlar)
        y.calistir(app, wb, "modGonderim.FormuTemizle")
        s.esit("Temizle düğmesi hata vermedi", "", y.son_mesaj(app, wb))
        kalan = [ad for ad in alanlar if str(ws.Range(ad).Value or "").strip()]
        s.kontrol("Temizle düğmesi bütün alanları boşalttı", not kalan, str(kalan))

        # Eksik zorunlu alanla gonderim reddedilmeli.
        _ekrana_yaz(ws, {"frm_ad_soyad": "Yalnız Ad"})
        y.calistir(app, wb, "modGonderim.OneriGonder")
        mesaj = y.son_mesaj(app, wb)
        s.kontrol("Eksik alanla gönderim reddedildi",
                  mesaj.startswith("hata:") and "Sicil No" in mesaj, mesaj)

    oneriler, _ = y.depo_oku(o.yonetim_kitap)
    s.esit("Form ekranından tam olarak bir kayıt yazıldı",
           onceki + 1, len(oneriler))
    kayit = oneriler[-1]
    s.esit("Formdaki değer kayda geçti", "Elif Karaduman", kayit["ad_soyad"])
    s.esit("Öneri başlığı kayda geçti", "Tek ekranda rapor",
           kayit["oneri_basligi"])
    s.esit("Numara sıradaki değeri aldı", f"PRJ-{y.yil()}-0004", kayit["oneri_no"])


# ==========================================================================
#  7. Yonetim kitabi ASLA kaydedilmez
#
#  Kitap salt okunur acilir ve ekranda yapilan her sey yalnizca bellektedir.
#  Bir kaydetme, bellekteki ESKI kopyayi diskin uzerine yazabilir ve o arada
#  personelin gonderdigi butun onerileri silebilirdi.
# ==========================================================================
def _kitap_kaydedilmez(s, app, o):
    print("  · yönetim kitabının gerçek açılışı (Workbook_Open) ve kaydetme yasağı")

    onceki_boyut = os.path.getsize(o.yonetim_kitap)
    onceki_oneri = len(y.depo_oku(o.yonetim_kitap)[0])
    s.esit("Bu ana kadar yedek alınmamıştı", [], o.yedek_dosyalar())

    # Olaylar acik: kitabin GERCEK acilis yolu boyle calisir. Diger testler
    # EnableEvents=False ile kostugu icin Workbook_Open'i yalnizca burasi
    # gorur -- ve orasi yedegi alip kilidi birakan yerdir.
    app.EnableEvents = True
    try:
        wb = app.Workbooks.Open(os.path.abspath(o.yonetim_kitap), 0, False, None,
                                y.DOSYA_SIFRESI, "", True)
        try:
            y.sessiz_mod(app, wb, True)

            s.kontrol("Açılışta kitap kendini salt okunura aldı (kilidi bıraktı)",
                      bool(wb.ReadOnly))
            s.esit("Açılışta günlük yedek alındı", 1, len(o.yedek_dosyalar()))
            s.kontrol("Açılışta yalnızca Giriş ekranı görünür",
                      all(w.Visible != -1 for w in wb.Worksheets
                          if w.Name != "Giriş"))

            wb.Worksheets("Veri").Cells(500, 1).Value = "bozma denemesi"
            try:
                wb.Save()
            except Exception:
                pass
            mesaj = y.son_mesaj(app, wb)
            s.kontrol("Kaydetme engellendi ve kullanıcıya açıklandı",
                      mesaj.startswith("bilgi:") and "kaydedilmez" in mesaj, mesaj)
        finally:
            wb.Saved = True
            wb.Close(SaveChanges=False)
    finally:
        app.EnableEvents = False

    s.esit("Disk dosyası değişmedi", onceki_boyut,
           os.path.getsize(o.yonetim_kitap))
    s.esit("Depodaki öneriler yerinde", onceki_oneri,
           len(y.depo_oku(o.yonetim_kitap)[0]))


# ==========================================================================
#  8. Gunluk yedek
# ==========================================================================
def _yedek(s, app, o):
    print("  · günlük yedek")

    # Yedek bir onceki testte, kitabin gercek acilisinda alindi.
    with y.kitap(app, o.yonetim_kitap, salt_okunur=True) as wb:
        yedekler = o.yedek_dosyalar()
        s.esit("Bir yedek kopya var", 1, len(yedekler))
        if yedekler:
            s.kontrol("Yedek adı tarihi taşıyor",
                      re.fullmatch(r"ProjeYonetim_\d{8}\.xlsm", yedekler[0])
                      is not None, yedekler[0])
            yol = os.path.join(o.yedek, yedekler[0])
            s.esit("Yedek, dosyanın birebir kopyası",
                   os.path.getsize(o.yonetim_kitap), os.path.getsize(yol))
            with open(yol, "rb") as f:
                s.kontrol("Yedek de şifreli",
                          f.read(8) == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")

        # Ayni gun ikinci cagri yeni kopya URETMEZ.
        y.calistir(app, wb, "modDepo.YedekAl")
        s.esit("Aynı gün ikinci yedek alınmadı", 1, len(o.yedek_dosyalar()))


# ==========================================================================
#  9. Depo baskasinda yazma kipindeyken gonderim
#
#  Sistemin en kritik yeni davranisi. Excel bu durumda HATA VERMEZ, dosyayi
#  sessizce SALT OKUNUR acar; modDepo bunu catisma sayip yeniden dener ve
#  butce dolunca anlasilir bir hata verir. Kullanicinin formu KAYBOLMAZ.
# ==========================================================================
def _kilitliyken_gonderim(s, app, o):
    print("  · depo başkasında yazma kipindeyken gönderim (≈20 sn)")

    with y.excel() as tutucu:
        kilitli = tutucu.Workbooks.Open(os.path.abspath(o.yonetim_kitap), 0,
                                        False, None, y.DOSYA_SIFRESI, "", True)
        try:
            s.kontrol("Tutucu kitabı yazma kipinde açtı", not kilitli.ReadOnly)

            baslangic = time.monotonic()
            with y.kitap(app, o.oneri_kitap, salt_okunur=True) as wb:
                sonuc = y.calistir(
                    app, wb, "modGonderim.TestGonderimi",
                    "Kilit Testi", "10100", "Depo kilitliyken gönderim denemesi.",
                    "Kilit denemesi", "Bir çözüm önerisi.", "Fayda")
            sure = time.monotonic() - baslangic

            s.kontrol("Kilitliyken gönderim anlaşılır bir hatayla döndü",
                      str(sonuc).startswith("HATA:")
                      and "kullanılıyor" in str(sonuc), str(sonuc)[:120])
            # Kilit hemen alinir (rakip yok); beklenen sure YAZMA butcesidir:
            # depo acilmayi denemeye devam eder ve butce dolunca pes eder.
            s.kontrol("Yazma bütçesi kadar bekledi, sonsuza kadar değil",
                      10 <= sure <= 70, f"{sure:.1f} sn")
        finally:
            kilitli.Saved = True
            kilitli.Close(SaveChanges=False)

    # Kilit birakilinca ayni gonderim calismali.
    with y.kitap(app, o.oneri_kitap, salt_okunur=True) as wb:
        no = y.calistir(
            app, wb, "modGonderim.TestGonderimi",
            "Kilit Testi", "10100", "Kilit bırakıldıktan sonra gönderim.",
            "Kilit sonrası", "Bir çözüm önerisi.", "Fayda")
    s.kontrol("Kilit bırakılınca gönderim yeniden çalışıyor",
              str(no).startswith("PRJ-"), str(no))


if __name__ == "__main__":
    sys.exit(calistir())
