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
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pywintypes
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
            _rapor(s, app, o)
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
               os.path.normcase(o.kaizen),
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
    s.kontrol("Öneri numarası kısa ve ON-YYMMDD-XXX biçiminde",
              o.no1.startswith("ON-") and len(o.no1) == 13
              and o.no1.count("-") == 2, o.no1)
    s.kontrol("Öneri numarasında karışabilen harf/rakam yok (0 O 1 I)",
              not set(o.no1.split("-")[2]) & set("O0I1"), o.no1)

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
        s.esit("Öncelik sınıfı hesaplanıp gösterildi", "Hızlı Kazanım",
               str(ws.Range("dg_oncelik").Value or ""))
        s.esit("Geçmiş listesinde en yeni değerlendirme üstte",
               "Pilot Uygulamada", str(ws.Cells(34, 4).Value or ""))

        # Reddetme gerekcesiz kabul edilmemeli.
        _ekrana_yaz(ws, {"dg_yeni_durum": "Reddedildi", "dg_not": ""})
        y.calistir(app, wb, "modDegerlendirme.DegerlendirmeKaydet")
        mesaj = y.son_mesaj(app, wb)
        s.kontrol("Gerekçesiz reddetme engellendi",
                  mesaj.startswith("hata:") and "gerekçe" in mesaj, mesaj)

        # Puanlanmadan kabul edilmemeli.
        _ekrana_yaz(ws, {"dg_yeni_durum": "Planlandı", "dg_etki": "",
                         "dg_efor": "", "dg_not": "Uygun."})
        y.calistir(app, wb, "modDegerlendirme.DegerlendirmeKaydet")
        mesaj = y.son_mesaj(app, wb)
        s.kontrol("Puanlanmadan kabul engellendi",
                  mesaj.startswith("hata:") and "etki ve efor" in mesaj, mesaj)

        # Gecerli bir kayit
        onceki = len(o.degerlendirme_dosyalar())
        _ekrana_yaz(ws, {
            "dg_yeni_durum": "Ölçümleniyor",
            "dg_etki": 4, "dg_efor": 2, "dg_saat": 150, "dg_tl": 36000,
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
        s.esit("Ekrandan kaydedilen saat kazanımı uygulandı", 150.0, veri["saat"])
        s.esit("Ekrandan kaydedilen TL tasarrufu uygulandı", 36000.0, veri["tl"])


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
                  "ON-" in str(ws.Range("frm_bant").Value or ""),
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
    yarim = os.path.join(o.oneriler_yil(), "ON-260101-ZZZ.txt")
    with open(yarim, "w", encoding="utf-8") as f:
        f.write("sema=3\noneri_no=ON-260101-ZZZ\ntarih=2026-01-01T00:00:00\n"
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
        s.kontrol("Konsol özeti yazıldı",
                  "öneri okundu" in str(
                      wb.Worksheets("Konsol").Range("knsl_ozet").Value or ""),
                  str(wb.Worksheets("Konsol").Range("knsl_ozet").Value))

        y.calistir(app, wb, "modKonsolide.KonsoluCiz")
        ws = wb.Worksheets("Konsol")
        s.esit("Konsol tablosuna üç satır yazıldı", 3,
               sum(1 for r in range(9, 20)
                   if str(ws.Cells(r, 2).Value or "").startswith("ON-")))

        s.esit("Yeni öneriler 'Yeni' durumuyla başlıyor", "Yeni",
               str(ws.Cells(9, 6).Value or ""))
        s.esit("Puanlanmamış etki boş gösteriliyor", "",
               str(ws.Cells(9, 7).Value or ""))

    os.remove(yarim)


# ==========================================================================
#  3. Degerlendirme (ekle-only)
# ==========================================================================
def _degerlendirme(s, app, o):
    print("  · değerlendirme ve geçmiş")

    with y.kitap(app, o.yonetim_kitap) as wb:
        # Birinci degerlendirme: puanlanir ve planlanir.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no1, "Planlandı", 4, 2, 120.0, 30000.0,
                   "Şube müdürüyle görüşüldü, uygulanabilir.")

        dosyalar_1 = o.degerlendirme_dosyalar()
        s.esit("İlk değerlendirme bir olay dosyası oluşturdu", 1, len(dosyalar_1))
        ilk_icerik = open(dosyalar_1[0], "rb").read()

        adet = y.calistir(app, wb, "modKonsolide.VeriyiKur")
        s.esit("Konsolidasyon öneri sayısını değiştirmedi", 3, adet)

        veri = _veri_satiri(wb, o.no1)
        s.esit("Durum olaydan türetildi", "Planlandı", veri["durum"])
        s.esit("Etki puanı uygulandı", 4, veri["etki"])
        s.esit("Efor puanı uygulandı", 2, veri["efor"])
        s.esit("Öncelik sınıfı hesaplandı (etki 4 / efor 2)",
               "Hızlı Kazanım", veri["oncelik"])

        # Ikinci degerlendirme: YALNIZCA durum degisir, puanlar verilmez.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no1, "Pilot Uygulamada", 0, 0, 0.0, 0.0,
                   "Pilot başladı.")

        dosyalar_2 = o.degerlendirme_dosyalar()
        s.esit("İkinci değerlendirme ikinci dosyayı ekledi", 2, len(dosyalar_2))
        s.kontrol("İlk olay dosyası olduğu gibi duruyor (ekle-only)",
                  os.path.exists(dosyalar_1[0])
                  and open(dosyalar_1[0], "rb").read() == ilk_icerik)

        y.calistir(app, wb, "modKonsolide.VeriyiKur")
        veri = _veri_satiri(wb, o.no1)
        s.esit("Son olay durumu belirledi", "Pilot Uygulamada", veri["durum"])
        s.esit("Kısmi olay önceki etki puanını SİLMEDİ", 4, veri["etki"])
        s.esit("Kısmi olay önceki efor puanını SİLMEDİ", 2, veri["efor"])
        s.esit("Kısmi olay önceki saat kazanımını SİLMEDİ", 120.0, veri["saat"])
        s.esit("Olay sayısı iki", 2, veri["olay_sayisi"])

        # Diger iki oneri: biri reddedilir, biri planlanir ama uygulanmaz.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no2, "Planlandı", 5, 4, 200.0, 50000.0,
                   "Kabul edildi, BT planına alındı.")
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no3, "Reddedildi", 2, 5, 0.0, 0.0,
                   "Mevzuat üç imzayı zorunlu kılıyor.")

        y.calistir(app, wb, "modKonsolide.VeriyiKur")
        s.esit("Yüksek etki / yüksek efor = Büyük Proje",
               "Büyük Proje", _veri_satiri(wb, o.no2)["oncelik"])
        s.esit("Düşük etki / yüksek efor = Değerlendirme Dışı",
               "Değerlendirme Dışı", _veri_satiri(wb, o.no3)["oncelik"])

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
                "etki": int(ws.Cells(r, 10).Value or 0),
                "efor": int(ws.Cells(r, 11).Value or 0),
                "oncelik": str(ws.Cells(r, 12).Value or ""),
                "saat": float(ws.Cells(r, 13).Value or 0),
                "tl": float(ws.Cells(r, 14).Value or 0),
                "olay_sayisi": int(ws.Cells(r, 19).Value or 0),
            }
    raise AssertionError(f"Veri sayfasında bulunamadı: {oneri_no}")


# ==========================================================================
#  4. Gostergeler -- tasarruf kurali
# ==========================================================================
def _gostergeler(s, app, o):
    print("  · göstergeler ve tasarruf kuralı")

    with y.kitap(app, o.yonetim_kitap) as wb:
        y.calistir(app, wb, "modKonsolide.VeriyiKur")
        ham = y.calistir(app, wb, "modPano.TestGostergeleri")
        g = dict(p.split("=", 1) for p in ham.split(";"))

        s.esit("Toplam öneri", "3", g["toplam"])
        s.esit("Bekleyen öneri yok (üçü de sonuçlandı)", "0", g["bekleyen"])
        s.esit("Uygulamaya geçmiş bir öneri", "1", g["uygulanan"])

        # KRITIK: 2 numarali oneri "Planlandi" -- 200 saat / 50.000 TL girmisti.
        # Planlandi henuz uygulanmis sayilmaz; toplama GIRMEMELIDIR.
        s.esit("Tasarruf yalnızca uygulanmış öneriden sayıldı (saat)",
               120.0, float(g["saat"]))
        s.esit("Tasarruf yalnızca uygulanmış öneriden sayıldı (TL)",
               30000.0, float(g["tl"]))

        s.esit("Hızlı kazanım sayısı", "1", g["hizli"])
        s.esit("Matris: Hızlı Kazanım", "1", g["mtx_hizli"])
        s.esit("Matris: Büyük Proje", "1", g["mtx_buyuk"])
        s.esit("Matris: Değerlendirme Dışı", "1", g["mtx_disi"])

        # Simdi 2 numarali oneriyi uygulamaya al: tasarruf artik sayilmali.
        y.calistir(app, wb, "modDegerlendirme.TestDegerlendirmesi",
                   o.no2, "Standartlaştırıldı", 0, 0, 0.0, 0.0,
                   "Tüm operasyona yaygınlaştırıldı.")
        y.calistir(app, wb, "modKonsolide.VeriyiKur")
        ham = y.calistir(app, wb, "modPano.TestGostergeleri")
        g = dict(p.split("=", 1) for p in ham.split(";"))

        s.esit("Uygulamaya geçince tasarruf toplama eklendi (saat)",
               320.0, float(g["saat"]))
        s.esit("Uygulamaya geçince tasarruf toplama eklendi (TL)",
               80000.0, float(g["tl"]))

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
#  5. Parolali rapor
# ==========================================================================
def _rapor(s, app, o):
    print("  · parolalı rapor")

    parola = "Kaizen!2026"
    with y.kitap(app, o.yonetim_kitap) as wb:
        y.calistir(app, wb, "modKonsolide.VeriyiKur")
        yol = y.calistir(app, wb, "modRapor.TestRaporu", "Tümü", parola)

    s.kontrol("Rapor dosyası üretildi", bool(yol) and os.path.exists(yol), yol)
    if not yol or not os.path.exists(yol):
        return

    s.kontrol("Rapor 'rapor' klasörüne yazıldı",
              os.path.normcase(os.path.dirname(yol)) == os.path.normcase(o.rapor))

    # Sifreli bir dosya OLE bilesik belgesidir; duz .xlsx (ZIP) degildir.
    with open(yol, "rb") as f:
        imza = f.read(8)
    s.kontrol("Rapor şifreli kapsayıcıya yazıldı (düz ZIP değil)",
              imza == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1", imza.hex())

    # Düz .xlsx bir ZIP arşividir; şifrelenmiş dosya ZIP olarak okunamaz.
    zip_olarak_okundu = True
    try:
        zipfile.ZipFile(yol)
    except zipfile.BadZipFile:
        zip_olarak_okundu = False
    s.kontrol("İçerik parolasız okunamıyor (ZIP olarak açılmıyor)",
              not zip_olarak_okundu)

    yanlis_parolayla_acildi = False
    try:
        wb2 = y.sifreli_ac(app, yol, "yanlis-parola")
        wb2.Close(SaveChanges=False)
        yanlis_parolayla_acildi = True
    except pywintypes.com_error:
        pass
    s.kontrol("Yanlış parolayla AÇILAMIYOR", not yanlis_parolayla_acildi)

    try:
        wb3 = y.sifreli_ac(app, yol, parola)
        sayfalar = {ws.Name for ws in wb3.Worksheets}
        satir_sayisi = wb3.Worksheets("Öneriler").UsedRange.Rows.Count
        ozet_baslik = str(wb3.Worksheets("Özet").Range("B2").Value or "")
        wb3.Close(SaveChanges=False)
        s.kontrol("Doğru parolayla açılıyor", True)
        s.esit("Rapor iki sayfadan oluşuyor", {"Özet", "Öneriler"}, sayfalar)
        s.esit("Rapor tablosunda başlık + üç öneri var", 4, satir_sayisi)
        s.esit("Rapor kapak başlığı yazıldı", "Kaizen Yönetim Raporu", ozet_baslik)
    except pywintypes.com_error as hata:
        s.kontrol("Doğru parolayla açılıyor", False, str(hata))


if __name__ == "__main__":
    sys.exit(calistir())
