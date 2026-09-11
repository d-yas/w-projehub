# -*- coding: utf-8 -*-
r"""Uretim dogrulamasi -- Excel acmadan, saniyeler icinde.

Dort soruyu yanitlar:
  1) Iki .xlsm gercekten uretildi mi ve icinde bir VBA projesi var mi?
  2) Yonetim kitabi gercekten SIFRELI mi ve bilinen parolayla aciliyor mu?
  3) Sayfa duzeni, adlandirilmis araliklar ve dogrulama listeleri yerinde mi?
  4) AYNI BILGININ IKI KOPYASI birbirinden ayrilmis mi?

Dorduncu madde bu testin asil degeri. Sistemde bilincli tekrar vardir:
  · renkler             -> tasarim.py       ile modTasarim.bas
  · tablo koordinatlari -> uret_yonetim.py  ile modKonsolide/modDegerlendirme
  · liste icerikleri    -> kur.py           ile modModel.bas
  · DEPO SUTUN DUZENI   -> uret_yonetim.py  ile modDepo.bas
Sonuncusu en tehlikelisidir: sutunlar sessizce kayarsa gonderim yanlis
sutuna yazilir ve hicbir sey hata vermez.

    python testler\test_uretim.py

NOT: Sifreli kitabin ICINE bakmak icin msoffcrypto-tool gerekir
(pip install msoffcrypto-tool). Kurulu degilse ic kontroller atlanir;
"sifreli mi" kontrolu her zaman calisir.
"""

import io
import os
import re
import sys
import zipfile

KOK = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(KOK, "kaynak"))
sys.path.insert(0, KOK)

import yardimci as y                                   # noqa: E402

CIKTI = os.path.join(KOK, "cikti")
VBA = os.path.join(KOK, "kaynak", "vba")

ONERI = os.path.join(CIKTI, "ProjeOneri.xlsm")
TAKIP = os.path.join(CIKTI, "ProjeTakip.xlsm")
YONETIM = os.path.join(CIKTI, "yonetim", "ProjeYonetim.xlsm")

OLE_IMZASI = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _bas_oku(ad):
    with open(os.path.join(VBA, ad), "r", encoding="utf-8") as f:
        return f.read()


def _vba_sabiti(kaynak, ad):
    """`Public Const AD As Long = 9` -> 9"""
    m = re.search(rf"Const\s+{ad}\s+As\s+\w+\s*=\s*(-?\d+)", kaynak)
    return int(m.group(1)) if m else None


def _vba_metin_sabitleri(kaynak):
    return {m.group(1): m.group(2) for m in
            re.finditer(r'Const\s+(\w+)\s+As\s+String\s*=\s*"([^"]*)"', kaynak)}


def _vba_dizi(kaynak, fonksiyon, sabitler):
    """Bir fonksiyonun icindeki Array(...) elemanlarini cozer.

    Elemanlar ya dogrudan dizge sabitidir ya da modulde tanimli bir Const'tur.
    """
    m = re.search(rf"Function\s+{fonksiyon}\s*\(.*?\n(.*?)End Function",
                  kaynak, re.DOTALL)
    if not m:
        return None
    govde = m.group(1).replace("_\n", " ")
    ic = re.search(r"Array\((.*?)\)", govde, re.DOTALL)
    if not ic:
        return None

    sonuc = []
    for parca in ic.group(1).split(","):
        parca = parca.strip()
        if not parca:
            continue
        if parca.startswith('"') and parca.endswith('"'):
            sonuc.append(parca[1:-1])
        elif parca in sabitler:
            sonuc.append(sabitler[parca])
        else:
            sonuc.append(parca)
    return sonuc


# VBA'da modul duzeyi degiskenler YALNIZCA bildirim bolumunde -- ilk
# Sub/Function'dan once -- durabilir. Yordamlarin arasina konan bir
# "Private x As Boolean" modulu DERLENMEZ; uretim ve dugme dogrulamasi bunu
# gormez, hata ancak o yordam ilk kez calistirildiginda ortaya cikar.
# Gorunmeyen bir Excel'de ise hic gorunmez: makro sessizce geri donmez.
_YORDAM = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?"
                     r"(?:Static\s+)?(?:Sub|Function|Property)\s", re.I)
_BILDIRIM = re.compile(r"^(?:Public|Private|Dim)\s+(?!Sub|Function|"
                       r"Property|Const|Type|Enum|Declare)"
                       r"(\w+)", re.I)


def _gec_kalmis_bildirimler(kaynak):
    """Ilk yordamdan SONRA gelen modul duzeyi degisken bildirimleri."""
    yordam_gorundu = False
    sorunlar = []
    for satir in kaynak.splitlines():
        if _YORDAM.match(satir):
            yordam_gorundu = True
        elif yordam_gorundu:
            e = _BILDIRIM.match(satir)
            if e:
                sorunlar.append(e.group(1))
    return sorunlar


def _sifreyi_coz(yol, parola):
    """Sifreli kitabi cozup BytesIO dondurur. msoffcrypto yoksa None."""
    try:
        import msoffcrypto
    except ImportError:
        return None
    with open(yol, "rb") as f:
        dosya = msoffcrypto.OfficeFile(f)
        dosya.load_key(password=parola)
        tampon = io.BytesIO()
        dosya.decrypt(tampon)
    tampon.seek(0)
    return tampon


def uret_yonetim_modulu():
    """uret_yonetim, KAYNAK sys.path'e eklendikten sonra import edilebilir."""
    import uret_yonetim
    return uret_yonetim


def calistir():
    s = y.Sonuc("Üretim")

    # ----------------------------------------------------------------------
    # 1. Dosyalar, sifreleme ve VBA projesi
    # ----------------------------------------------------------------------
    print("  · üretilen dosyalar ve şifreleme")
    kitaplar = (("ProjeOneri.xlsm", ONERI), ("ProjeTakip.xlsm", TAKIP),
                ("ProjeYonetim.xlsm", YONETIM))
    for ad, yol in kitaplar:
        s.kontrol(f"{ad} üretildi", os.path.exists(yol), yol)

    if not all(os.path.exists(yol) for _, yol in kitaplar):
        print("\n  Önce: python kur.py")
        return s.bitir()

    # Personel kitaplari parolasizdir; ikisi de acilir acilmaz calismali.
    for ad, yol in (("ProjeOneri.xlsm", ONERI), ("ProjeTakip.xlsm", TAKIP)):
        s.kontrol(f"{ad} şifresiz (personel açabilmeli)", zipfile.is_zipfile(yol))
        with zipfile.ZipFile(yol) as z:
            s.kontrol(f"{ad} bir VBA projesi içeriyor",
                      "xl/vbaProject.bin" in set(z.namelist()))

    # Yonetim kitabi VERI DEPOSUDUR: gizliligin siniri dosya parolasidir.
    # Sifreli bir OOXML dosyasi zip degil, OLE bilesik dosyadir.
    with open(YONETIM, "rb") as f:
        bas = f.read(8)
    s.kontrol("ProjeYonetim.xlsm şifreli (OLE imzası)", bas == OLE_IMZASI,
              repr(bas))
    s.kontrol("ProjeYonetim.xlsm düz zip DEĞİL", not zipfile.is_zipfile(YONETIM))

    tampon = _sifreyi_coz(YONETIM, y.DOSYA_SIFRESI)
    if tampon is None:
        print("    msoffcrypto-tool kurulu değil — kitabın içine bakan "
              "kontroller atlanıyor")
    else:
        s.kontrol("ProjeYonetim.xlsm modAyar'daki parolayla açılıyor", True)
        with zipfile.ZipFile(tampon) as z:
            s.kontrol("ProjeYonetim.xlsm bir VBA projesi içeriyor",
                      "xl/vbaProject.bin" in set(z.namelist()))

    # ----------------------------------------------------------------------
    # 2. Sayfa duzeni ve adlandirilmis araliklar
    # ----------------------------------------------------------------------
    print("  · sayfa düzeni ve adlandırılmış aralıklar")
    from openpyxl import load_workbook

    wb1 = load_workbook(ONERI, keep_vba=True)
    s.esit("Öneri kitabı üç sayfadan oluşuyor",
           {"Giriş", "Öneri Formu", "Listeler"}, set(wb1.sheetnames))
    s.esit("Giriş dışındaki ekran gizli", "veryHidden",
           wb1["Öneri Formu"].sheet_state)
    s.esit("Giriş ekranı görünür", "visible", wb1["Giriş"].sheet_state)

    form_adlari = {"frm_bant", "frm_ad_soyad", "frm_sicil_no",
                   "frm_mevcut", "frm_baslik", "frm_cozum", "frm_fayda"}
    s.kontrol("Form alanlarının tamamı adlandırılmış aralık",
              form_adlari <= set(wb1.defined_names),
              str(sorted(form_adlari - set(wb1.defined_names))))

    ws = wb1["Öneri Formu"]
    s.kontrol("Form sayfasında kılavuz çizgileri kapalı",
              not ws.sheet_view.showGridLines)
    s.kontrol("Form giriş alanları kilitsiz (koruma altında yazılabilir)",
              ws["C10"].protection.locked is False)
    s.kontrol("Etiket hücreleri kilitli", ws["B10"].protection.locked is not False)
    s.kontrol("Form sayfasında açılır liste kalmadı (birim ve israf kaldırıldı)",
              not ws.data_validations.dataValidation,
              str([dv.formula1 for dv in ws.data_validations.dataValidation]))

    # --- Takip kitabi: tek ekran, VERI YOK -------------------------------
    # Kitap parolasizdir; bir depo sayfasi buraya girerse herkesin onerisi
    # herkesin elinde olur. Sorgu depoyu gizli ornekte okur (modTakip).
    wb3 = load_workbook(TAKIP, keep_vba=True)
    s.esit("Takip kitabı tek sayfadan oluşuyor", ["Takip"], wb3.sheetnames)
    s.esit("Takip ekranı görünür", "visible", wb3["Takip"].sheet_state)

    takip_adlari = {"tk_bant", "tk_no", "tk_sicil", "tk_sonuc_no", "tk_tarih",
                    "tk_baslik", "tk_durum", "tk_guncelleme"}
    s.kontrol("Takip alanlarının tamamı adlandırılmış aralık",
              takip_adlari <= set(wb3.defined_names),
              str(sorted(takip_adlari - set(wb3.defined_names))))

    def _takip_hucresi(ad):
        _, koordinat = next(iter(wb3.defined_names[ad].destinations))
        return wb3["Takip"][koordinat.replace("$", "")]

    if takip_adlari <= set(wb3.defined_names):
        for ad in ("tk_no", "tk_sicil"):
            h = _takip_hucresi(ad)
            s.kontrol(f"Takip girişi kilitsiz: {ad}", h.protection.locked is False)
            s.esit(f"Takip girişi metin biçimli (baştaki sıfır korunur): {ad}",
                   "@", h.number_format)
        for ad in ("tk_sonuc_no", "tk_durum", "tk_guncelleme"):
            s.kontrol(f"Takip sonucu kilitli (elle yazılamaz): {ad}",
                      _takip_hucresi(ad).protection.locked is not False)

    import uret_takip
    wt = wb3["Takip"]
    cf_takip = {str(alan.sqref): len(kurallar)
                for alan, kurallar in wt.conditional_formatting._cf_rules.items()}
    gecmis_durum = (f"C{uret_takip.GECMIS_ILK}:"
                    f"C{uret_takip.GECMIS_ILK + uret_takip.GECMIS_ADET - 1}")
    s.esit("Takip: güncel durumda sekiz durumun rengi var", 8, cf_takip.get("C25"))
    s.esit("Takip: geçmiş tablosunda sekiz durumun rengi var", 8,
           cf_takip.get(gecmis_durum))

    uy = uret_yonetim_modulu()

    if tampon is not None:
        tampon.seek(0)
        wb2 = load_workbook(tampon, keep_vba=True)
        s.esit("Yönetim kitabı dokuz sayfadan oluşuyor",
               {"Giriş", "Liste", "Değerlendirme", "Pano",
                "Oneriler", "Olaylar", "Veri", "PanoVeri", "Listeler"},
               set(wb2.sheetnames))
        s.esit("Liste kitaptaki ilk sayfa", "Liste", wb2.sheetnames[0])
        s.kontrol("Çalışma ve depo sayfaları girişten önce gizli",
                  all(wb2[a].sheet_state == "veryHidden"
                      for a in ("Liste", "Değerlendirme", "Pano",
                                "Oneriler", "Olaylar", "Veri", "PanoVeri")))
        s.esit("Kitap giriş ekranı etkinken kaydedildi", "Giriş", wb2.active.title)

        # --- Depo sayfalarinin basliklari --------------------------------
        s.esit("Oneriler sayfasının başlıkları",
               uy.ONERILER_BASLIKLARI,
               [c.value for c in wb2["Oneriler"][1]][:len(uy.ONERILER_BASLIKLARI)])
        s.esit("Olaylar sayfasının başlıkları",
               uy.OLAYLAR_BASLIKLARI,
               [c.value for c in wb2["Olaylar"][1]][:len(uy.OLAYLAR_BASLIKLARI)])
        s.kontrol("Depo sayfaları boş üretiliyor (veri sonradan eklenir)",
                  all(wb2[a].cell(row=2, column=2).value in (None, "")
                      for a in ("Oneriler", "Olaylar")))

        pano_adlari = {"pano_toplam", "pano_bekleyen", "pano_uygulanan",
                       "pano_bu_ay", "pano_guncelleme"}
        deg_adlari = {"dg_bant", "dg_oneri_no", "dg_tarih", "dg_gonderen",
                      "dg_baslik", "dg_mevcut", "dg_cozum", "dg_fayda",
                      "dg_yeni_durum", "dg_not"}
        diger = {"liste_ozet", "veri_adet", "giris_bant"}
        eksik = (pano_adlari | deg_adlari | diger) - set(wb2.defined_names)
        s.kontrol("Yönetim kitabının tüm adlandırılmış aralıkları yerinde",
                  not eksik, str(sorted(eksik)))

        liste = wb2["Liste"]
        s.esit("Listede başlık satırı ve ilk sütunlar donduruldu", "C9",
               liste.freeze_panes)
        cf = {str(alan.sqref): len(kurallar)
              for alan, kurallar in liste.conditional_formatting._cf_rules.items()}
        s.esit("Durum sütununda sekiz durumun sekiz rengi var", 8, cf.get("F9:F2000"))

    # Panoda dort KPI karti tek sirada durur; bos kart yuvasi kalmamalidir.
    s.esit("Panoda dört KPI kartı tanımlı", 4, len(uy.PANO_KART_SUTUNLARI))
    s.esit("Her KPI kartının bir göstergesi var", 4, len(uy.PANO_KARTLARI))

    # ----------------------------------------------------------------------
    # 3. Kopyalanan bilgi tutarli mi?
    # ----------------------------------------------------------------------
    print("  · kaynaklar arası tutarlılık")

    import kur
    import tasarim
    import uret_yonetim

    s.esit("Renkler: tasarim.py ile modTasarim.bas aynı (35 renk)", 35,
           kur.tasarim_tutarliligini_dogrula())

    konsolide = _bas_oku("modKonsolide.bas")
    for py_deger, vba_ad, aciklama in (
        (uret_yonetim.LISTE_ILK_SATIR, "LISTE_ILK_SATIR", "tablonun ilk satırı"),
        (uret_yonetim.LISTE_ILK_SUTUN, "LISTE_ILK_SUTUN", "tablonun ilk sütunu"),
        (uret_yonetim.LISTE_SON_SUTUN, "LISTE_SON_SUTUN", "tablonun son sütunu"),
    ):
        s.esit(f"Liste koordinatı — {aciklama}", py_deger,
               _vba_sabiti(konsolide, vba_ad))

    # Oturum LISTE ile acilir (ekibin gunluk isi); Pano bir dugme uzaktadir.
    giris = re.search(r"Public Sub SistemeGir\b(.*?)\nEnd Sub", konsolide, re.DOTALL)
    s.kontrol("Yönetim oturumu Liste ekranıyla açılıyor",
              giris is not None and re.search(
                  r"OturumAc\s+Array\(.*?\),\s*modUI\.SAYFA_LISTE",
                  giris.group(1), re.DOTALL) is not None)

    degerlendirme = _bas_oku("modDegerlendirme.bas")
    s.esit("Geçmiş tablosunun ilk satırı", uret_yonetim.DEG_GECMIS_ILK,
           _vba_sabiti(degerlendirme, "GECMIS_ILK_SATIR"))
    s.esit("Geçmiş tablosunun satır sayısı", uret_yonetim.DEG_GECMIS_ADET,
           _vba_sabiti(degerlendirme, "GECMIS_AZAMI"))

    model = _bas_oku("modModel.bas")
    listeler = kur.listeler()
    s.esit("Doğrulama listesi 'durum' modModel.bas ile aynı",
           listeler["durum"],
           _vba_dizi(model, "Durumlar", _vba_metin_sabitleri(model)))

    # Takip ekrani her durumun anlamini yazar; eksik bir durum orada bos
    # bir hucre olarak gorunur.
    aciklama = re.search(r"Function DurumAciklamasi\b(.*?)End Function",
                         model, re.DOTALL)
    model_sabitleri = _vba_metin_sabitleri(model)
    aciklanan = {model_sabitleri.get(ad) for ad in
                 re.findall(r"Case\s+(DURUM_\w+)\s*:", aciklama.group(1))} \
        if aciklama else set()
    s.esit("Her durumun takip ekranında bir açıklaması var",
           set(listeler["durum"]), aciklanan)

    takip = _bas_oku("modTakip.bas")
    s.esit("Takip geçmiş tablosunun ilk satırı", uret_takip.GECMIS_ILK,
           _vba_sabiti(takip, "GECMIS_ILK_SATIR"))
    s.esit("Takip geçmiş tablosunun satır sayısı", uret_takip.GECMIS_ADET,
           _vba_sabiti(takip, "GECMIS_AZAMI"))
    s.esit("Takip sayfa adı — modUI.SAYFA_TAKIP", uret_takip.SAYFA_TAKIP,
           _vba_metin_sabitleri(_bas_oku("modUI.bas")).get("SAYFA_TAKIP"))

    # Veri sayfasinin sutun sayisi ile modKonsolide'nin beklentisi
    s.esit("Veri sayfasının sütun sayısı", len(uret_yonetim.VERI_BASLIKLARI),
           _vba_sabiti(konsolide, "V_SUTUN_SAYISI"))

    # ----------------------------------------------------------------------
    # 4. DEPO SUTUN DUZENI -- en kritik kopya
    # ----------------------------------------------------------------------
    print("  · depo sütun düzeni (uret_yonetim.py ↔ modDepo.bas)")
    depo = _bas_oku("modDepo.bas")
    depo_metin = _vba_metin_sabitleri(depo)

    s.esit("Depo sayfa adı — Oneriler", uret_yonetim.SAYFA_ONERILER,
           depo_metin.get("SAYFA_ONERILER"))
    s.esit("Depo sayfa adı — Olaylar", uret_yonetim.SAYFA_OLAYLAR,
           depo_metin.get("SAYFA_OLAYLAR"))

    for onek, basliklar, sayi_sabiti in (
        ("O_", uret_yonetim.ONERILER_BASLIKLARI, "O_SUTUN_SAYISI"),
        ("E_", uret_yonetim.OLAYLAR_BASLIKLARI, "E_SUTUN_SAYISI"),
    ):
        s.esit(f"{onek}SUTUN_SAYISI başlık sayısıyla aynı", len(basliklar),
               _vba_sabiti(depo, sayi_sabiti))
        for konum, baslik in enumerate(basliklar, start=1):
            sabit = onek + baslik.upper()
            s.esit(f"Sütun {sabit}", konum, _vba_sabiti(depo, sabit))

    # Alan sirasi da ayni: sozlukten satira cevirirken bu diziler kullanilir.
    s.esit("modDepo.OneriAlanlari başlıklarla aynı sırada",
           uret_yonetim.ONERILER_BASLIKLARI,
           _vba_dizi(depo, "OneriAlanlari", depo_metin))
    s.esit("modDepo.OlayAlanlari başlıklarla aynı sırada",
           uret_yonetim.OLAYLAR_BASLIKLARI,
           _vba_dizi(depo, "OlayAlanlari", depo_metin))

    # --- Ayar sabitleri: testler ve kur.py bunlari .bas'tan okur ----------
    ayar = _bas_oku("modAyar.bas")
    ayar_metin = _vba_metin_sabitleri(ayar)
    s.esit("Şema sürümü", "4", ayar_metin.get("SEMA_SURUMU"))
    s.esit("Yönetim kitabının adı", "ProjeYonetim.xlsm",
           ayar_metin.get("DOSYA_YONETIM"))
    s.esit("Yedek klasörünün adı", "yedek", ayar_metin.get("KLASOR_YEDEK"))
    s.kontrol("Yedek sayısı makul", 2 <= (_vba_sabiti(ayar, "YEDEK_ADET") or 0) <= 60,
              str(_vba_sabiti(ayar, "YEDEK_ADET")))
    sifre = ayar_metin.get("SIFRE_DOSYA", "")
    s.kontrol("Depo parolası tanımlı ve boş değil", len(sifre) >= 6, repr(sifre))
    s.kontrol("Depo parolası makul uzunlukta (≤ 24)", len(sifre) <= 24, repr(sifre))
    s.kontrol("Öneri numarası sıralı biçimde üretiliyor",
              re.search(r'Format\$\(enBuyuk \+ 1, "0000"\)', depo) is not None)

    # ----------------------------------------------------------------------
    # 5. VBA kaynaklari
    # ----------------------------------------------------------------------
    print("  · VBA kaynak dosyaları")
    bas_dosyalari = [d for d in os.listdir(VBA) if d.endswith(".bas")]
    s.kontrol("On bir VBA modülü var", len(bas_dosyalari) == 11,
              str(sorted(bas_dosyalari)))

    tanimli = set()
    for tanim in (kur.oneri_tanimi(), kur.yonetim_tanimi(), kur.takip_tanimi()):
        tanimli |= {os.path.basename(m) for m in tanim["moduller"]}
    s.kontrol("Her VBA modülü en az bir kitaba giriyor",
              set(bas_dosyalari) <= tanimli,
              str(sorted(set(bas_dosyalari) - tanimli)))

    for d in sorted(bas_dosyalari):
        kaynak = _bas_oku(d)
        ad = os.path.splitext(d)[0]
        s.kontrol(f"{d} modül adını bildiriyor",
                  re.search(rf'Attribute VB_Name = "{ad}"', kaynak) is not None)
        s.kontrol(f"{d} Option Explicit kullanıyor", "Option Explicit" in kaynak)
        gec = _gec_kalmis_bildirimler(kaynak)
        s.kontrol(f"{d} modül değişkenleri bildirim bölümünde", not gec,
                  ", ".join(gec))

    # Testlerin BAGIMLI oldugu giris noktalari. Bunlardan biri kaybolursa
    # her uctan uca test AZAMI_SURE kadar donup "makro yanıt vermiyor" der;
    # sebebi ise eksik yordam yuzunden modulun derlenmemesidir. Burada adiyla
    # aranmalari, o teshisi saniyelere indirir.
    ui = _bas_oku("modUI.bas")
    for yordam in ("SessizModAyarla", "SessizMi", "SonMesaj", "MesajKaydet",
                   "OturumAc", "OturumKapat", "KorumalariKur"):
        s.kontrol(f"modUI.bas {yordam} tanımlıyor",
                  re.search(rf"(Sub|Function)\s+{yordam}\s*\(", ui) is not None)

    for yordam in ("DepoAc", "KitapKapat", "GizliExcelAc", "GizliExcelKapat",
                   "OneriEkle", "OlayEkle", "OlayEkleVeCek", "DepoyuCek",
                   "SonSatir", "TabloOku", "AralikDizisi", "OneriNoUret",
                   "SaltOkunuraGec", "SaltOkunurMu", "YedekAl",
                   "KilidiAl", "KilidiBirak", "BayatKilidiTemizle",
                   "KaydetmeyiDene", "DepoOneDriveAltindaMi",
                   "OneDriveAciklamasi", "TestDepoSayilari",
                   "OneriSatirlariniOku"):
        s.kontrol(f"modDepo.bas {yordam} tanımlıyor",
                  re.search(rf"(Sub|Function)\s+{yordam}\s*\(", depo) is not None)

    # Yazma yollari KILIT ALTINDA olmali. Kilit cagrisi bir yeniden duzenlemede
    # kaybolursa hicbir sey hata vermez; yalnizca iki kisi ayni anda gonderim
    # yaptiginda bir satir sessizce kaybolur. Bu yuzden adiyla aranir.
    for kapi in ("OneriEkle", "OlayEkle", "OlayEkleVeCek"):
        desen = r"Public Function " + kapi + r"\b(.*?)\nEnd Function"
        govde = re.search(desen, depo, re.DOTALL)
        s.kontrol(f"modDepo.{kapi} yazmadan önce kilidi alıyor",
                  govde is not None and "KilidiAl" in govde.group(1))
        s.kontrol(f"modDepo.{kapi} kilidi bırakıyor",
                  govde is not None and "KilidiBirak" in govde.group(1))

    # modDepo kullaniciyla KONUSMAZ: gorunmez bir Excel'de acilan bir MsgBox
    # makroyu sonsuza kadar bekletir. (Yorum satirlari elenir; modulun kendi
    # aciklamasi bu kurali ANLATIR ve aramada yanlis eslesme yapardi.)
    depo_kod = "\n".join(satir for satir in depo.splitlines()
                         if not satir.lstrip().startswith("'"))
    s.kontrol("modDepo.bas içinde MsgBox/InputBox yok",
              not re.search(r"\b(MsgBox|InputBox)\b", depo_kod))

    # OneDrive konumu, ekip tarafinda YAZMADAN ONCE denetlenmeli. Bu kontrol
    # kaybolursa hata yine olur ama kullanici yirmi saniye bekleyip yaniltici
    # bir sebep okur -- uretimde tam olarak boyle bulundu.
    for kapi in ("OlayEkle", "OlayEkleVeCek"):
        desen = r"Public Function " + kapi + r"\b(.*?)\nEnd Function"
        govde = re.search(desen, depo, re.DOTALL)
        s.kontrol(f"modDepo.{kapi} OneDrive konumunu önden denetliyor",
                  govde is not None and "DepoOneDriveAltindaMi" in govde.group(1))

    # Dosya tabanli veri katmani tumden kalkti; kalintisi kalmamali.
    dosyaio = _bas_oku("modDosyaIO.bas")
    olu = [ad for ad in ("UTF8Yaz", "UTF8Oku", "TamKayitYaz", "KayitOku",
                         "KayitMetniUret", "DosyalariTara", "KisaRastgele",
                         "ZamanDamgasi", "DegerKacisla")
           if re.search(rf"(Sub|Function)\s+{ad}\s*\(", dosyaio)]
    s.kontrol("modDosyaIO'da dosya tabanlı kayıt yordamı kalmadı", not olu,
              ", ".join(olu))
    s.kontrol("modDosyaIO.bas OneDriveAltindaMi tanımlıyor",
              re.search(r"Function\s+OneDriveAltindaMi\s*\(", dosyaio)
              is not None)

    # ThisWorkbook_Yonetim: kitabin kendisi asla kaydedilmemeli.
    twy = os.path.join(VBA, "ThisWorkbook_Yonetim.vba")
    with open(twy, "r", encoding="utf-8") as f:
        yonetim_kod = f.read()
    s.kontrol("Yönetim kitabında kaydetme engelli (Workbook_BeforeSave)",
              re.search(r"Workbook_BeforeSave.*\n(?:.*\n)*?\s*Cancel = True",
                        yonetim_kod) is not None)
    s.kontrol("Açılışta yedek alınıyor", "modDepo.YedekAl" in yonetim_kod)
    s.kontrol("Açılışta salt okunura geçiliyor",
              "modDepo.SaltOkunuraGec" in yonetim_kod)
    s.kontrol("Açılışta OneDrive konumu uyarılıyor",
              "DepoOneDriveAltindaMi" in yonetim_kod)

    # --- Takip ekrani: gizlilik kurallari KODDA da durmali -----------------
    # Ekibin karar notlari ve degerlendirenin kimligi oneri sahibine
    # gosterilmez (bilincli karar). Bir yeniden duzenlemede bu sutunlardan
    # biri ekrana eklenirse hicbir sey hata vermez; bu yuzden adiyla aranir.
    takip_kod = "\n".join(satir for satir in takip.splitlines()
                          if not satir.lstrip().startswith("'"))
    for sutun in ("E_KARAR_NOTU", "E_DEGERLENDIREN_KULLANICI",
                  "E_DEGERLENDIREN_BILGISAYAR"):
        s.kontrol(f"modTakip {sutun} sütununu okumuyor (öneri sahibine gösterilmez)",
                  sutun not in takip_kod)
    s.kontrol("modTakip sorguda sicil numarasını denetliyor",
              "O_SICIL_NO" in takip_kod)
    s.kontrol("modTakip depoyu kitaba kopyalamıyor (DepoyuCek yok)",
              "DepoyuCek" not in takip_kod)
    s.kontrol("Liste ve takip ekranı durumu aynı kuraldan alıyor",
              "modModel.OlaySonrasiDurum" in takip_kod
              and "modModel.OlaySonrasiDurum" in konsolide)

    with open(os.path.join(VBA, "ThisWorkbook_Takip.vba"), "r", encoding="utf-8") as f:
        s.kontrol("Takip kitabı açılışta ekranı sıfırlıyor",
                  "modTakip.EkraniSifirla" in f.read())

    # --- VBA kaynagi ANSI kod sayfasinda saklanir --------------------------
    # Turkce harfler cp1254'te vardir; "✓" ve "⚠" gibi isaretler yoktur. Kod
    # satirina duz yazildiklarinda hucreye "?" olarak iner ve hicbir sey hata
    # vermez (bantlarda boyle goruldu). Bu karakterler modTasarim.IsaretOnay /
    # IsaretUyari ile uretilir. Yorum satirlari zararsizdir, elenir.
    tasan = []
    for d in sorted(os.listdir(VBA)):
        if not d.endswith((".bas", ".vba")):
            continue
        for no, satir in enumerate(_bas_oku(d).splitlines(), start=1):
            if satir.lstrip().startswith("'"):
                continue
            if any(not ch.encode("cp1254", "ignore") for ch in satir):
                tasan.append(f"{d}:{no}")
    s.kontrol("VBA kod satırlarında cp1254 dışı karakter yok (hücreye '?' iner)",
              not tasan, ", ".join(tasan))

    return s.bitir()


if __name__ == "__main__":
    sys.exit(calistir())
