# -*- coding: utf-8 -*-
r"""Uretim dogrulamasi -- Excel acmadan, saniyeler icinde.

Uc soruyu yanitlar:
  1) Iki .xlsm gercekten uretildi mi ve icinde bir VBA projesi var mi?
  2) Sayfa duzeni, adlandirilmis araliklar ve dogrulama listeleri yerinde mi?
  3) AYNI BILGININ IKI KOPYASI birbirinden ayrilmis mi?

Ucuncu madde bu testin asil degeri. Sistemde uc yerde bilincli tekrar var:
  · renkler        -> tasarim.py  ile modTasarim.bas
  · tablo koordinatlari -> uret_yonetim.py ile modKonsolide/modDegerlendirme
  · liste icerikleri    -> kur.py (hucre dogrulamasi) ile modModel.bas (calisma zamani)
Bu kopyalarin sessizce birbirinden ayrilmasi, uretimde fark edilmesi zor
hatalara yol acar. Test onlari her calistirmada karsilastirir.

    python testler\test_uretim.py
"""

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
YONETIM = os.path.join(CIKTI, "yonetim", "ProjeYonetim.xlsm")


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
_BILDIRIM = re.compile(r"^(?:Public|Private|Dim)\s+(?!Sub|Function|"
                       r"Property|Const|Type|Enum|Declare)"
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


def uret_yonetim_modulu():
    """uret_yonetim, KAYNAK sys.path'e eklendikten sonra import edilebilir."""
    import uret_yonetim
    return uret_yonetim


def calistir():
    s = y.Sonuc("Üretim")

    # ----------------------------------------------------------------------
    # 1. Dosyalar ve VBA projesi
    # ----------------------------------------------------------------------
    print("  · üretilen dosyalar")
    for ad, yol in (("ProjeOneri.xlsm", ONERI), ("ProjeYonetim.xlsm", YONETIM)):
        if not s.kontrol(f"{ad} üretildi", os.path.exists(yol), yol):
            continue
        with zipfile.ZipFile(yol) as z:
            adlar = set(z.namelist())
        s.kontrol(f"{ad} bir VBA projesi içeriyor", "xl/vbaProject.bin" in adlar)

    if not (os.path.exists(ONERI) and os.path.exists(YONETIM)):
        print("\n  Önce: python kur.py")
        return s.bitir()

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

    wb2 = load_workbook(YONETIM, keep_vba=True)
    s.esit("Yönetim kitabı yedi sayfadan oluşuyor",
           {"Giriş", "Liste", "Değerlendirme", "Pano",
            "Veri", "PanoVeri", "Listeler"}, set(wb2.sheetnames))
    s.esit("Pano kitaptaki ilk sayfa", "Pano", wb2.sheetnames[0])
    s.kontrol("Çalışma sayfaları girişten önce gizli",
              all(wb2[a].sheet_state == "veryHidden"
                  for a in ("Liste", "Değerlendirme", "Pano",
                            "Veri", "PanoVeri")))
    s.esit("Kitap giriş ekranı etkinken kaydedildi", "Giriş",
           wb2.active.title)

    pano_adlari = {"pano_toplam", "pano_bekleyen", "pano_uygulanan",
                   "pano_bu_ay", "pano_guncelleme"}
    deg_adlari = {"dg_bant", "dg_oneri_no", "dg_tarih", "dg_gonderen",
                  "dg_baslik", "dg_mevcut", "dg_cozum", "dg_fayda",
                  "dg_yeni_durum", "dg_not"}
    diger = {"liste_ozet", "veri_adet"}
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
    s.esit("Panoda dört KPI kartı tanımlı", 4,
           len(uret_yonetim_modulu().PANO_KART_SUTUNLARI))
    s.esit("Her KPI kartının bir göstergesi var", 4,
           len(uret_yonetim_modulu().PANO_KARTLARI))

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

    degerlendirme = _bas_oku("modDegerlendirme.bas")
    s.esit("Geçmiş tablosunun ilk satırı", uret_yonetim.DEG_GECMIS_ILK,
           _vba_sabiti(degerlendirme, "GECMIS_ILK_SATIR"))
    s.esit("Geçmiş tablosunun satır sayısı", uret_yonetim.DEG_GECMIS_ADET,
           _vba_sabiti(degerlendirme, "GECMIS_AZAMI"))

    model = _bas_oku("modModel.bas")
    sabitler = _vba_metin_sabitleri(model)
    listeler = kur.listeler()
    s.esit("Doğrulama listesi 'durum' modModel.bas ile aynı",
           listeler["durum"], _vba_dizi(model, "Durumlar", sabitler))

    # Veri sayfasinin sutun sayisi ile modKonsolide'nin beklentisi
    s.esit("Veri sayfasının sütun sayısı", len(uret_yonetim.VERI_BASLIKLARI),
           _vba_sabiti(konsolide, "V_SUTUN_SAYISI"))

    # ----------------------------------------------------------------------
    # 4. VBA kaynaklari
    # ----------------------------------------------------------------------
    print("  · VBA kaynak dosyaları")
    bas_dosyalari = [d for d in os.listdir(VBA) if d.endswith(".bas")]
    s.kontrol("Dokuz VBA modülü var", len(bas_dosyalari) == 9,
              str(len(bas_dosyalari)))
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
    # her uctan uca test 90 saniye donup "makro yanıt vermiyor" der; sebebi
    # ise eksik yordam yuzunden modulun derlenmemesidir. Burada adiyla
    # aranmalari, o teshisi saniyelere indirir.
    ui = _bas_oku("modUI.bas")
    for yordam in ("SessizModAyarla", "SessizMi", "SonMesaj", "MesajKaydet",
                   "OturumAc", "OturumKapat", "KorumalariKur"):
        s.kontrol(f"modUI.bas {yordam} tanımlıyor",
                  re.search(rf"(Sub|Function)\s+{yordam}\s*\(", ui) is not None)

    return s.bitir()


if __name__ == "__main__":
    sys.exit(calistir())
