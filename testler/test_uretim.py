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

ONERI = os.path.join(CIKTI, "KaizenOneri.xlsm")
YONETIM = os.path.join(CIKTI, "yonetim", "KaizenYonetim.xlsm")


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


def calistir():
    s = y.Sonuc("Üretim")

    # ----------------------------------------------------------------------
    # 1. Dosyalar ve VBA projesi
    # ----------------------------------------------------------------------
    print("  · üretilen dosyalar")
    for ad, yol in (("KaizenOneri.xlsm", ONERI), ("KaizenYonetim.xlsm", YONETIM)):
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
    s.esit("Yönetim kitabı sekiz sayfadan oluşuyor",
           {"Giriş", "Konsol", "Değerlendirme", "Pano", "Rapor",
            "Veri", "PanoVeri", "Listeler"}, set(wb2.sheetnames))
    s.kontrol("Çalışma sayfaları girişten önce gizli",
              all(wb2[a].sheet_state == "veryHidden"
                  for a in ("Konsol", "Değerlendirme", "Pano", "Rapor",
                            "Veri", "PanoVeri")))

    pano_adlari = {"pano_toplam", "pano_bekleyen", "pano_uygulanan", "pano_saat",
                   "pano_tl", "pano_hizli", "pano_kabul_orani",
                   "pano_yanit_suresi", "pano_bu_ay", "pano_mtx_hizli",
                   "pano_mtx_buyuk", "pano_mtx_doldurma", "pano_mtx_disi",
                   "pano_guncelleme"}
    deg_adlari = {"dg_bant", "dg_oneri_no", "dg_tarih", "dg_gonderen",
                  "dg_baslik", "dg_mevcut", "dg_cozum", "dg_fayda",
                  "dg_yeni_durum", "dg_etki", "dg_efor",
                  "dg_oncelik", "dg_saat", "dg_tl", "dg_not"}
    diger = {"knsl_ozet", "veri_adet", "rpr_donem", "rpr_son"}
    eksik = (pano_adlari | deg_adlari | diger) - set(wb2.defined_names)
    s.kontrol("Yönetim kitabının tüm adlandırılmış aralıkları yerinde",
              not eksik, str(sorted(eksik)))

    konsol = wb2["Konsol"]
    s.esit("Konsolda başlık satırı ve ilk sütunlar donduruldu", "C9",
           konsol.freeze_panes)
    cf = {str(alan.sqref): len(kurallar)
          for alan, kurallar in konsol.conditional_formatting._cf_rules.items()}
    s.esit("Durum sütununda sekiz durumun sekiz rengi var", 8, cf.get("F9:F2000"))
    s.esit("Öncelik sütununda dört sınıfın dört rengi var", 4, cf.get("I9:I2000"))

    # ----------------------------------------------------------------------
    # 3. Kopyalanan bilgi tutarli mi?
    # ----------------------------------------------------------------------
    print("  · kaynaklar arası tutarlılık")

    import kur
    import tasarim
    import uret_yonetim

    s.kontrol("Renkler: tasarim.py ile modTasarim.bas aynı",
              kur.tasarim_tutarliligini_dogrula() > 0)

    konsolide = _bas_oku("modKonsolide.bas")
    for py_deger, vba_ad, aciklama in (
        (uret_yonetim.KONSOL_ILK_SATIR, "KONSOL_ILK_SATIR", "tablonun ilk satırı"),
        (uret_yonetim.KONSOL_ILK_SUTUN, "KONSOL_ILK_SUTUN", "tablonun ilk sütunu"),
        (uret_yonetim.KONSOL_SON_SUTUN, "KONSOL_SON_SUTUN", "tablonun son sütunu"),
    ):
        s.esit(f"Konsol koordinatı — {aciklama}", py_deger,
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

    # Esik degerleri belgede anlatilanla ayni mi?
    ayar = _bas_oku("modAyar.bas")
    s.esit("Yüksek etki eşiği 3", 3, _vba_sabiti(ayar, "ESIK_YUKSEK_ETKI"))
    s.esit("Düşük efor eşiği 2", 2, _vba_sabiti(ayar, "ESIK_DUSUK_EFOR"))

    # ----------------------------------------------------------------------
    # 4. VBA kaynaklari
    # ----------------------------------------------------------------------
    print("  · VBA kaynak dosyaları")
    bas_dosyalari = [d for d in os.listdir(VBA) if d.endswith(".bas")]
    s.kontrol("On VBA modülü var", len(bas_dosyalari) == 10, str(len(bas_dosyalari)))
    for d in sorted(bas_dosyalari):
        kaynak = _bas_oku(d)
        ad = os.path.splitext(d)[0]
        s.kontrol(f"{d} modül adını bildiriyor",
                  re.search(rf'Attribute VB_Name = "{ad}"', kaynak) is not None)
        s.kontrol(f"{d} Option Explicit kullanıyor", "Option Explicit" in kaynak)

    return s.bitir()


if __name__ == "__main__":
    sys.exit(calistir())
