# -*- coding: utf-8 -*-
r"""Proje Öneri Sistemi -- üretim betiği.

    python kur.py            iki kitabı da üretir
    python kur.py oneri      yalnızca ProjeOneri.xlsm
    python kur.py yonetim    yalnızca ProjeYonetim.xlsm

Çalışma kitapları elle hazırlanmaz; kaynak kod düz metin dosyalarında durur
(kaynak\vba\*.bas ve kaynak\uret_*.py), kitaplar buradan sıfırdan üretilir.
Bir değişiklik yapmak için ilgili kaynak dosya düzenlenir ve bu betik yeniden
çalıştırılır. Böylece tasarım ve kod sürüm takibine girer, karşılaştırılabilir
ve gözden geçirilebilir olur.
"""

import os
import re
import shutil
import sys

KOK = os.path.dirname(os.path.abspath(__file__))
KAYNAK = os.path.join(KOK, "kaynak")
VBA = os.path.join(KAYNAK, "vba")
CIKTI = os.path.join(KOK, "cikti")

sys.path.insert(0, KAYNAK)

import com_kurulum as com          # noqa: E402
import tasarim                     # noqa: E402
import uret_ortak as u             # noqa: E402
from tasarim import RENK           # noqa: E402


# ==========================================================================
#  1. Tasarim tutarliligi
# ==========================================================================
_HEX_DESENI = re.compile(r'"([0-9A-Fa-f]{6})"\s*\'#HEX\s+(.+?)\s*$', re.MULTILINE)


def tasarim_tutarliligini_dogrula():
    """tasarim.py ile modTasarim.bas ayni renkleri mi soyluyor?

    Iki kaynak arasinda sessiz bir kayma, panonun bir yerde lacivert bir
    yerde baska bir mavi kullanmasina yol acardi. Uretim, uyusmazlikta durur.
    """
    bas_yolu = os.path.join(VBA, "modTasarim.bas")
    with open(bas_yolu, "r", encoding="utf-8") as f:
        icerik = f.read()

    vba_hexler = {ad: kod.upper() for kod, ad in _HEX_DESENI.findall(icerik)}
    py_hexler = {ad: kod.upper() for ad, kod in tasarim.tum_hexler().items()}

    sorunlar = []
    for ad, kod in py_hexler.items():
        if ad not in vba_hexler:
            sorunlar.append(f"modTasarim.bas içinde '#HEX {ad} etiketi yok")
        elif vba_hexler[ad] != kod:
            sorunlar.append(f"{ad}: tasarim.py={kod} ≠ modTasarim.bas={vba_hexler[ad]}")
    for ad in vba_hexler:
        if ad not in py_hexler:
            sorunlar.append(f"tasarim.py içinde tanımsız renk: {ad}")

    if sorunlar:
        raise SystemExit("Tasarım tutarlılık hatası:\n  - " + "\n  - ".join(sorunlar))
    return len(py_hexler)


# ==========================================================================
#  2. Ortak veri dogrulama listeleri
# ==========================================================================
def listeler():
    """Veri dogrulama listeleri -- modModel.bas ile ayni icerik.

    Not: modModel.bas calisma zamaninin kaynagidir; buradaki kopya yalnizca
    hucre dogrulama listelerini besler. Ikisi ayni sirayla tutulur.
    """
    return {
        "durum": ["Yeni", "Değerlendirmede", "Planlandı", "Pilot Uygulamada",
                  "Ölçümleniyor", "Standartlaştırıldı", "Beklemede", "Reddedildi"],
    }


def _bas(*adlar):
    return [os.path.join(VBA, ad + ".bas") for ad in adlar]


def _thisworkbook(ad):
    with open(os.path.join(VBA, ad), "r", encoding="utf-8") as f:
        return f.read()


# ==========================================================================
#  3. Kitap tanimlari
# ==========================================================================
def oneri_tanimi():
    import uret_oneri

    taslak = os.path.join(CIKTI, "_taslak_oneri.xlsx")
    hedef = os.path.join(CIKTI, "ProjeOneri.xlsm")
    return {
        "ad": "ProjeOneri.xlsm",
        "taslak": taslak,
        "hedef": hedef,
        "uret": lambda: uret_oneri.kitap_uret(taslak, listeler()),
        "moduller": _bas("modTasarim", "modAyar", "modDosyaIO", "modModel",
                         "modUI", "modGonderim"),
        "thisworkbook": _thisworkbook("ThisWorkbook_Oneri.vba"),
        # Dugme konumlari sayfa duzeniyle birlikte degistigi icin adresler
        # uret_oneri'den okunur; iki yerde ayri ayri tutulmaz.
        "dugmeler": [
            {"sayfa": uret_oneri.SAYFA_GIRIS,
             "hucre": uret_oneri.DUGME_YERLERI["giris"],
             "metin": "Sisteme Gir  →", "makro": "SistemeGir",
             "varyant": "birincil"},
            {"sayfa": uret_oneri.SAYFA_FORM,
             "hucre": uret_oneri.DUGME_YERLERI["gonder"],
             "metin": "Öneriyi Gönder", "makro": "OneriGonder",
             "varyant": "birincil"},
            {"sayfa": uret_oneri.SAYFA_FORM,
             "hucre": uret_oneri.DUGME_YERLERI["temizle"],
             "metin": "Formu Temizle", "makro": "FormuTemizle",
             "varyant": "ikincil",
             "sol_kaydir": uret_oneri.DUGME_KAYDIR["temizle"]},
            {"sayfa": uret_oneri.SAYFA_FORM,
             "hucre": uret_oneri.DUGME_YERLERI["cikis"],
             "metin": "Çıkış", "makro": "Cikis", "varyant": "sessiz",
             "genislik": 78.0, "yukseklik": 26.0},
        ],
        "ek_islem": None,
    }


def yonetim_tanimi():
    import uret_yonetim

    taslak = os.path.join(CIKTI, "_taslak_yonetim.xlsx")
    hedef = os.path.join(CIKTI, "yonetim", "ProjeYonetim.xlsm")
    return {
        "ad": "ProjeYonetim.xlsm",
        "taslak": taslak,
        "hedef": hedef,
        "uret": lambda: uret_yonetim.kitap_uret(taslak, listeler()),
        "moduller": _bas("modTasarim", "modAyar", "modDosyaIO", "modModel",
                         "modUI", "modKonsolide", "modDegerlendirme",
                         "modPano"),
        "thisworkbook": _thisworkbook("ThisWorkbook_Yonetim.vba"),
        "dugmeler": uret_yonetim.DUGMELER,
        "ek_islem": uret_yonetim.com_ek_islem,
    }


TANIMLAR = {"oneri": oneri_tanimi, "yonetim": yonetim_tanimi}


# ==========================================================================
#  4. Uretim
# ==========================================================================
def uret(secilenler):
    print("Proje Öneri Sistemi — üretim")
    print("=" * 62)

    sayi = tasarim_tutarliligini_dogrula()
    print(f"  [1/4] Tasarım tutarlılığı        {sayi} renk eşleşti")

    os.makedirs(CIKTI, exist_ok=True)

    tanimlar = [TANIMLAR[a]() for a in secilenler]

    for t in tanimlar:
        t["uret"]()
        print(f"  [2/4] Taslak üretildi            {os.path.basename(t['taslak'])}")

    with com.vba_erisimi_acik():
        with com.excel_oturumu() as app:
            for t in tanimlar:
                modul_adlari = com.kitap_isle(
                    app, t["taslak"], t["hedef"], t["moduller"],
                    t["thisworkbook"], t["dugmeler"],
                    ek_islem=t["ek_islem"],
                    koruma_sifresi="po-koruma",
                )
                print(f"  [3/4] VBA + düğme eklendi        {t['ad']}  "
                      f"({len(modul_adlari)} modül, {len(t['dugmeler'])} düğme)")

            for t in tanimlar:
                sorunlar = com.kitap_dogrula(
                    app, t["hedef"],
                    [os.path.splitext(os.path.basename(m))[0] for m in t["moduller"]],
                    [d["makro"] for d in t["dugmeler"]],
                )
                if sorunlar:
                    print(f"  [4/4] DOĞRULAMA BAŞARISIZ        {t['ad']}")
                    for s in sorunlar:
                        print(f"         ! {s}")
                    raise SystemExit(1)
                print(f"  [4/4] Doğrulandı                 {t['ad']}")

    for t in tanimlar:
        if os.path.exists(t["taslak"]):
            os.remove(t["taslak"])

    print("=" * 62)
    for t in tanimlar:
        print(f"  ✓ {os.path.relpath(t['hedef'], KOK)}")
    print()
    print("  Kurulum için KURULUM.md dosyasına bakın.")
    print("  Uyarı: çıktıları OneDrive ile eşlenmiş bir klasörden çalıştırmayın.")


def main():
    argumanlar = [a.lower() for a in sys.argv[1:] if not a.startswith("-")]
    if not argumanlar:
        secilenler = ["oneri", "yonetim"]
    else:
        bilinmeyen = [a for a in argumanlar if a not in TANIMLAR]
        if bilinmeyen:
            raise SystemExit(f"Bilinmeyen hedef: {', '.join(bilinmeyen)}  "
                             f"(geçerli: {', '.join(TANIMLAR)})")
        secilenler = argumanlar
    uret(secilenler)


if __name__ == "__main__":
    main()
