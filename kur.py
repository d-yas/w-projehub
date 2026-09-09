# -*- coding: utf-8 -*-
r"""Proje Öneri Sistemi -- üretim betiği.

    python kur.py            iki kitabı da üretir
    python kur.py oneri      yalnızca ProjeOneri.xlsm
    python kur.py yonetim    yalnızca ProjeYonetim.xlsm

    python kur.py yonetim --veri "\\sunucu\...\yonetim\ProjeYonetim.xlsm"
        Yönetim kitabını yeniden üretir AMA verdiğiniz kitaptaki önerileri ve
        değerlendirme geçmişini yenisine taşır. Yönetim kitabı veri deposu
        olduğu için, dolu bir kurulumu güncellerken bu şarttır -- aksi halde
        yeni dosya boş gelir.

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

# Cikti bir dosyaya ya da boruya yonlendirildiginde Python konsol degil
# YEREL kod sayfasini (Turkce Windows'ta cp1254) kullanir ve buradaki "✓"
# gibi karakterler UnicodeEncodeError verir. Betik o zaman, isini bitirmis
# olmasina ragmen hata koduyla biter.
for _akis in (sys.stdout, sys.stderr):
    try:
        _akis.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, ValueError):
        pass

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


def ayar_sabiti(ad):
    """modAyar.bas'taki bir String sabitini okur.

    Parolalar TEK yerde durur: modAyar.bas. Python tarafinda ikinci bir kopya
    tutulsaydi ikisi sessizce ayrilabilir ve uretilen kitap, kodun bekledigi
    paroladan baskasiyla sifrelenebilirdi.
    """
    with open(os.path.join(VBA, "modAyar.bas"), "r", encoding="utf-8") as f:
        icerik = f.read()
    eslesme = re.search(rf'Const\s+{ad}\s+As\s+String\s*=\s*"([^"]*)"', icerik)
    if not eslesme:
        raise SystemExit(f"modAyar.bas içinde {ad} sabiti bulunamadı.")
    return eslesme.group(1)


def ayar_sayisi(ad):
    """modAyar.bas'taki bir sayisal sabiti okur."""
    with open(os.path.join(VBA, "modAyar.bas"), "r", encoding="utf-8") as f:
        icerik = f.read()
    eslesme = re.search(rf"Const\s+{ad}\s+As\s+\w+\s*=\s*(-?\d+)", icerik)
    if not eslesme:
        raise SystemExit(f"modAyar.bas içinde {ad} sabiti bulunamadı.")
    return int(eslesme.group(1))


def _bas(*adlar):
    return [os.path.join(VBA, ad + ".bas") for ad in adlar]


def _thisworkbook(ad):
    with open(os.path.join(VBA, ad), "r", encoding="utf-8") as f:
        return f.read()


# ==========================================================================
#  3. Kitap tanimlari
# ==========================================================================
def oneri_tanimi(veri_kaynagi=None):
    import uret_oneri

    taslak = os.path.join(CIKTI, "_taslak_oneri.xlsx")
    hedef = os.path.join(CIKTI, "ProjeOneri.xlsm")
    return {
        "ad": "ProjeOneri.xlsm",
        "taslak": taslak,
        "hedef": hedef,
        "uret": lambda: uret_oneri.kitap_uret(taslak, listeler()),
        # modDepo iki kitapta da bulunur: gonderim tarafi da depoya yazar.
        "moduller": _bas("modTasarim", "modAyar", "modDosyaIO", "modModel",
                         "modUI", "modDepo", "modGonderim"),
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
        "dosya_sifresi": None,
        "veri_kaynagi": None,
        "veri_sayfalari": (),
    }


def yonetim_tanimi(veri_kaynagi=None):
    import uret_yonetim

    taslak = os.path.join(CIKTI, "_taslak_yonetim.xlsx")
    hedef = os.path.join(CIKTI, "yonetim", "ProjeYonetim.xlsm")
    return {
        "ad": "ProjeYonetim.xlsm",
        "taslak": taslak,
        "hedef": hedef,
        "uret": lambda: uret_yonetim.kitap_uret(taslak, listeler()),
        "moduller": _bas("modTasarim", "modAyar", "modDosyaIO", "modModel",
                         "modUI", "modDepo", "modKonsolide",
                         "modDegerlendirme", "modPano"),
        "thisworkbook": _thisworkbook("ThisWorkbook_Yonetim.vba"),
        "dugmeler": uret_yonetim.DUGMELER,
        "ek_islem": uret_yonetim.com_ek_islem,
        # Bu kitap veri deposudur, o yuzden acilis parolasiyla sifrelenir ve
        # yeniden uretilirken eski verisi tasinabilir.
        "dosya_sifresi": ayar_sabiti("SIFRE_DOSYA"),
        "veri_kaynagi": veri_kaynagi,
        "veri_sayfalari": (uret_yonetim.SAYFA_ONERILER,
                           uret_yonetim.SAYFA_OLAYLAR),
    }


TANIMLAR = {"oneri": oneri_tanimi, "yonetim": yonetim_tanimi}


# ==========================================================================
#  4. Uretim
# ==========================================================================
def uret(secilenler, veri_kaynagi=None):
    print("Proje Öneri Sistemi — üretim")
    print("=" * 62)

    sayi = tasarim_tutarliligini_dogrula()
    print(f"  [1/4] Tasarım tutarlılığı        {sayi} renk eşleşti")

    os.makedirs(CIKTI, exist_ok=True)

    tanimlar = [TANIMLAR[a](veri_kaynagi) for a in secilenler]

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
                    koruma_sifresi=ayar_sabiti("SIFRE_KORUMA"),
                    dosya_sifresi=t["dosya_sifresi"],
                    veri_kaynagi=t["veri_kaynagi"],
                    veri_sayfalari=t["veri_sayfalari"],
                )
                print(f"  [3/4] VBA + düğme eklendi        {t['ad']}  "
                      f"({len(modul_adlari)} modül, {len(t['dugmeler'])} düğme)"
                      + ("  [parolalı]" if t["dosya_sifresi"] else ""))

            for t in tanimlar:
                sorunlar = com.kitap_dogrula(
                    app, t["hedef"],
                    [os.path.splitext(os.path.basename(m))[0] for m in t["moduller"]],
                    [d["makro"] for d in t["dugmeler"]],
                    dosya_sifresi=t["dosya_sifresi"],
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
    print("  Kurulum için TALIMATNAME.md dosyasına bakın.")
    print("  Uyarı: çıktıları OneDrive ile eşlenmiş bir klasörden çalıştırmayın.")


def main():
    ham = sys.argv[1:]
    veri_kaynagi = None

    if "--veri" in ham:
        i = ham.index("--veri")
        if i + 1 >= len(ham):
            raise SystemExit("--veri bir dosya yolu bekler.")
        veri_kaynagi = ham[i + 1]
        if not os.path.isfile(veri_kaynagi):
            raise SystemExit(f"--veri için verilen dosya yok: {veri_kaynagi}")
        ham = ham[:i] + ham[i + 2:]

    argumanlar = [a.lower() for a in ham if not a.startswith("-")]
    if not argumanlar:
        secilenler = ["oneri", "yonetim"]
    else:
        bilinmeyen = [a for a in argumanlar if a not in TANIMLAR]
        if bilinmeyen:
            raise SystemExit(f"Bilinmeyen hedef: {', '.join(bilinmeyen)}  "
                             f"(geçerli: {', '.join(TANIMLAR)})")
        secilenler = argumanlar

    if veri_kaynagi and "yonetim" not in secilenler:
        raise SystemExit("--veri yalnızca 'yonetim' hedefiyle anlamlıdır.")

    uret(secilenler, veri_kaynagi)


if __name__ == "__main__":
    main()
