# -*- coding: utf-8 -*-
r"""Uctan uca dogrulama -- gercek Excel, gercek makrolar, gercek dosyalar.

Bu bir taklit (mock) testi degildir: uretilen .xlsm dosyalarini Excel'de acar,
VBA makrolarini calistirir ve sonuclari dosya sisteminden dogrular.

TESTIN GOREMEDIGI (bilincli sinir):
Makrolar COM uzerinden cagrilir. Bu yol Excel'in makro guvenlik ayarini,
"Icerigi Etkinlestir" uyarisini ve dugmelerin gorsel yerlesimini HIC gormez.
Yani burasi "kod dogru mu" sorusunu yanitlar, "kullanici bu dosyayi actiginda
ne olur" sorusunu yanitlamaz. Kurulumdan sonra KURULUM.md'deki elle kontrol
listesi mutlaka bir kez uygulanmalidir.

    python testler\test_uctan_uca.py
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import yardimci as y


def calistir():
    s = y.Sonuc("Uçtan uca")

    with y.ortam() as o:
        with y.excel() as app:
            # Sira onemli: sayi bekleyen testler once calisir, veri ekleyen
            # ekran testleri en sona birakilir.
            _gonderim(s, app, o)
            _konsolidasyon(s, app, o)
            _degerlendirme(s, app, o)
            _gostergeler(s, app, o)
            _degerlendirme_ekrani(s, app, o)
            _form_ekrani(s, app, o)

    return s.bitir()


# ==========================================================================
#  1. Gonderim
# ==========================================================================
def _gonderim(s, app, o):
    print("  · gönderim akışı")
    with y.kitap(app, o.oneri_kitap) as wb:
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

    dosyalar = o.oneri_dosyalar()
    s.esit("Üç gönderim üç ayrı dosya oluşturdu", 3, len(dosyalar))

    s.kontrol("Öneri numaraları benzersiz", len({o.no1, o.no2, o.no3}) == 3)
    s.kontrol("Öneri numarası kısa ve PRJ-YYXXX biçiminde",
              o.no1.startswith("PRJ-") and len(o.no1) == 9
              and o.no1.count("-") == 1, o.no1)
    s.kontrol("Öneri numarasında karışabilen harf/rakam yok (0 O 1 I)",
              not set(o.no1.split("-")[1][2:]) & set("O0I1"), o.no1)

    s.kontrol("Dosya adı öneri numarasıyla aynı",
              sorted(os.path.splitext(os.path.basename(d))[0] for d in dosyalar)
              == sorted([o.no1, o.no2, o.no3]))

    s.kontrol("Kayıtlar yıl alt klasörüne yazıldı",
              all(os.path.dirname(d) == o.oneriler_yil() for d in dosyalar))

    ilk = [d for d in dosyalar
           if os.path.basename(d).startswith(o.no1)][0]
    kayit = y.kayit_oku(ilk)

    s.kontrol("Dosya UTF-8 BOM'suz", not y.bom_var_mi(ilk))
    s.kontrol("Kayıt sonu işareti son satırda",
              open(ilk, encoding="utf-8").read().rstrip("\n").endswith("kayit_sonu=1"))
    s.esit("Türkçe karakterler kayıtta bozulmadı", "Ayşe Çağlar", kayit["ad_soyad"])
    s.esit("Şema sürümü yazıldı", "3", kayit["sema"])
    s.kontrol("Çok satırlı alan geri çözüldü",
              kayit["mevcut_durum"].count("\n") == 1, repr(kayit["mevcut_durum"]))
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
#  6. Degerlendirme ekraninin GERCEK dugme yolu
#
#  Yukaridaki degerlendirme testi kayit yazma yolunu dogrular. Burasi
#  kullanicinin izledigi yolu dener: oneriyi ekrana yukle, alanlari doldur,
#  "Kaydet". Bu yol daha once birlesik hucre hatasi yuzunden yarida
#  kaliyordu ve hata gorunmuyordu.
# ==========================================================================
def _degerlendirme_ekrani(s, app, o):
    print("  · değerlendirme ekranı (Yükle / Kaydet düğmeleri)")

    with y.kitap(app, o.yonetim_kitap) as wb:
        y.calistir(app, wb, "modKonsolide.VeriyiKur")
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
        onceki = len(o.degerlendirme_dosyalar())
        _ekrana_yaz(ws, {
            "dg_yeni_durum": "Ölçümleniyor",
            "dg_not": "Pilot sonuçları ölçülmeye başlandı."})
        y.calistir(app, wb, "modDegerlendirme.DegerlendirmeKaydet")
        app.EnableEvents = False        # Kaydet sonrasi yenileme olaylari geri acar

        s.esit("Kaydet düğmesi yeni olay dosyası ekledi",
               onceki + 1, len(o.degerlendirme_dosyalar()))
        s.kontrol("Kayıt sonrası onay bandı yazıldı",
                  "kaydedildi" in str(ws.Range("dg_bant").Value or ""),
                  str(ws.Range("dg_bant").Value))

        y.calistir(app, wb, "modKonsolide.VeriyiKur")
        veri = _veri_satiri(wb, o.no1)
        s.esit("Ekrandan kaydedilen durum uygulandı", "Ölçümleniyor", veri["durum"])
        s.esit("Ekrandan kaydedilen karar notu uygulandı",
               "Pilot sonuçları ölçülmeye başlandı.", veri["karar_notu"])


def _ekrana_yaz(ws, degerler):
    """Korumalı sayfadaki giriş alanlarını doldurur."""
    y.korumasiz(ws)
    for ad, deger in degerler.items():
        ws.Range(ad).Value = deger


# ==========================================================================
#  7. Form ekraninin GERCEK dugme yolu
#
#  Yukaridaki testler kayit yazma yolunu dogrular. Burasi kullanicinin
#  gerçekte izledigi yolu dener: hucreler doldurulur, "Gönder" makrosu
#  calistirilir, sonra "Temizle". Bu yol daha once sessizce bozulmustu --
#  birlesik hucrelerde ClearContents calismiyordu ve hata gizleniyordu.
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

    onceki_dosyalar = set(o.oneri_dosyalar())
    onceki = len(onceki_dosyalar)
    with y.kitap(app, o.oneri_kitap) as wb:
        ws = wb.Worksheets("Öneri Formu")
        ws.Visible = -1
        _ekrana_yaz(ws, alanlar)

        y.calistir(app, wb, "modGonderim.OneriGonder")
        mesaj = y.son_mesaj(app, wb)

        s.kontrol("Gönder düğmesi başarı mesajı verdi",
                  mesaj.startswith("bilgi:") and "Öneri numaranız" in mesaj, mesaj)
        s.kontrol("Onay bandı öneri numarasını gösteriyor",
                  "PRJ-" in str(ws.Range("frm_bant").Value or ""),
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

    # Öneri numarası artık zaman sıralı değil; yeni dosya ada göre değil
    # kümeler farkıyla bulunur.
    yeni_dosyalar = set(o.oneri_dosyalar()) - onceki_dosyalar
    s.esit("Form ekranından tam olarak bir kayıt yazıldı", 1, len(yeni_dosyalar))
    if not yeni_dosyalar:
        return

    kayit = y.kayit_oku(next(iter(yeni_dosyalar)))
    s.esit("Formdaki değer kayda geçti", "Elif Karaduman", kayit["ad_soyad"])
    s.esit("Öneri başlığı kayda geçti", "Tek ekranda rapor", kayit["oneri_basligi"])


# ==========================================================================
#  2. Konsolidasyon
# ==========================================================================
def _konsolidasyon(s, app, o):
    print("  · konsolidasyon")

    # Yarida kesilmis bir yazma taklidi: "kayit_sonu" satiri yok.
    yarim = os.path.join(o.oneriler_yil(), "PRJ-26ZZZ.txt")
    with open(yarim, "w", encoding="utf-8") as f:
        f.write("sema=3\noneri_no=PRJ-26ZZZ\ntarih=2026-01-01T00:00:00\n"
                "ad_soyad=Yarım Kayıt\n")

    with y.kitap(app, o.yonetim_kitap) as wb:
        s.kontrol("Yönetim kitabının ThisWorkbook modülü derleniyor",
                  y.derleme_sinamasi(app, wb))

        adet = y.calistir(app, wb, "modKonsolide.VeriyiKur")
        s.esit("Yarım yazılmış kayıt okunmadı", 3, adet)

        # Gonderimler tek ve kalici klasordedir; okunmak onlari HICBIR YERE
        # tasimaz. Yarim kayit da yerinde birakilir.
        s.esit("Okuma dosyaları yerinden oynatmadı", 4, len(o.oneri_dosyalar()))
        s.kontrol("Kayıtlar yönetim klasörünün altında",
                  all(d.startswith(o.yonetim) for d in o.oneri_dosyalar()))
        s.kontrol("Yarım kayıt yerinde duruyor", os.path.exists(yarim))

        s.esit("Yeniden okuma aynı sonucu veriyor", 3,
               y.calistir(app, wb, "modKonsolide.VeriyiKur"))

        # Dugmelerin cagirdigi tam yol: olaylari geri actigi icin ThisWorkbook'u
        # da devreye sokar. Daha once burada sessizce kilitleniyordu.
        y.calistir(app, wb, "modKonsolide.OnerileriYenile")
        app.EnableEvents = False
        s.esit("“Önerileri Yenile” düğmesi hatasız çalıştı", "",
               y.son_mesaj(app, wb))
        s.kontrol("Liste özeti yazıldı",
                  "öneri okundu" in str(
                      wb.Worksheets("Liste").Range("liste_ozet").Value or ""),
                  str(wb.Worksheets("Liste").Range("liste_ozet").Value))

        y.calistir(app, wb, "modKonsolide.ListeyiCiz")
        ws = wb.Worksheets("Liste")
        s.esit("Liste tablosuna üç satır yazıldı", 3,
               sum(1 for r in range(9, 20)
                   if str(ws.Cells(r, 2).Value or "").startswith("PRJ-")))

        s.esit("Yeni öneriler 'Yeni' durumuyla başlıyor", "Yeni",
               str(ws.Cells(9, 6).Value or ""))

    os.remove(yarim)


# ==========================================================================
#  3. Degerlendirme (ekle-only)
# ==========================================================================
def _degerlendirme(s, app, o):
    print("  · değerlendirme ve geçmiş")

    with y.kitap(app, o.yonetim_kitap) as wb:
        # Birinci degerlendirme: planlanir.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no1, "Planlandı",
                   "Şube müdürüyle görüşüldü, uygulanabilir.")

        dosyalar_1 = o.degerlendirme_dosyalar()
        s.esit("İlk değerlendirme bir olay dosyası oluşturdu", 1, len(dosyalar_1))
        ilk_icerik = open(dosyalar_1[0], "rb").read()

        adet = y.calistir(app, wb, "modKonsolide.VeriyiKur")
        s.esit("Konsolidasyon öneri sayısını değiştirmedi", 3, adet)

        veri = _veri_satiri(wb, o.no1)
        s.esit("Durum olaydan türetildi", "Planlandı", veri["durum"])
        s.esit("Karar notu uygulandı", "Şube müdürüyle görüşüldü, uygulanabilir.",
               veri["karar_notu"])

        # Ikinci degerlendirme: YALNIZCA durum degisir, not verilmez.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no1, "Pilot Uygulamada", "")

        dosyalar_2 = o.degerlendirme_dosyalar()
        s.esit("İkinci değerlendirme ikinci dosyayı ekledi", 2, len(dosyalar_2))
        s.kontrol("İlk olay dosyası olduğu gibi duruyor (ekle-only)",
                  os.path.exists(dosyalar_1[0])
                  and open(dosyalar_1[0], "rb").read() == ilk_icerik)

        y.calistir(app, wb, "modKonsolide.VeriyiKur")
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

        y.calistir(app, wb, "modKonsolide.VeriyiKur")
        s.esit("İkinci öneri planlandı olarak işlendi",
               "Planlandı", _veri_satiri(wb, o.no2)["durum"])
        s.esit("Üçüncü öneri reddedildi olarak işlendi",
               "Reddedildi", _veri_satiri(wb, o.no3)["durum"])

        # Gecmis: bir onerinin tum olaylari
        gecmis = y.calistir(app, wb, "modDegerlendirme.OlayDosyalari", o.no1)
        s.esit("Bir önerinin geçmişi iki olaydan oluşuyor", 2, gecmis.Count())


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

    with y.kitap(app, o.yonetim_kitap) as wb:
        y.calistir(app, wb, "modKonsolide.VeriyiKur")
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
        y.calistir(app, wb, "modKonsolide.VeriyiKur")
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


if __name__ == "__main__":
    sys.exit(calistir())
