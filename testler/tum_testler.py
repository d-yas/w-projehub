# -*- coding: utf-8 -*-
r"""Butun testleri sirayla calistirir.

    python testler\tum_testler.py

Sira ucuzdan pahaliya dogrudur: yapisal kontroller Excel acmadan saniyeler
icinde biter, digerleri gercek Excel ornekleri baslatir ve dakikalar surer
(her gonderim depoyu acip kaydeden kisa bir islemdir). Bir asama basarisiz
olursa sonrakiler yine de calisir; boylece tek kosuda butun sorunlar gorulur.
"""

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import test_eszamanlilik
import test_izinler
import test_uctan_uca
import test_uretim
import yardimci as y


def main():
    asamalar = [
        ("Üretim", test_uretim.calistir),
        ("Uçtan uca", test_uctan_uca.calistir),
        ("Eşzamanlılık", test_eszamanlilik.calistir),
        ("İzinler", test_izinler.calistir),
    ]

    # Onceki bir "python kur.py" cagrisindan kalan Excel hala kapaniyor
    # olabilir; sizinti kontrollerinin temeli oturana kadar beklenir.
    baslangic = y.excel_sayisi_bekle(0, 30)
    if baslangic:
        print(f"UYARI: {baslangic} adet EXCEL.EXE çalışıyor ve kapanmadı. "
              f"Sızıntı kontrolleri bu sayıyı temel alır.")

    kalan = []
    t0 = time.monotonic()
    for ad, calistir in asamalar:
        print(f"\n### {ad}")
        if calistir() != 0:
            kalan.append(ad)

    print(f"\nToplam süre: {time.monotonic() - t0:.0f} sn")

    artik = y.excel_sayisi_bekle(baslangic)
    if artik > baslangic:
        print(f"UYARI: arkada {artik - baslangic} adet görünmez EXCEL.EXE kaldı.")
        print("       taskkill /F /IM EXCEL.EXE")

    print()
    if kalan:
        print(f"BAŞARISIZ aşamalar: {', '.join(kalan)}")
        return 1
    print("Tüm aşamalar geçti.")
    print()
    print("Hatırlatma: bu testler makroları COM üzerinden çağırır; Excel'in")
    print("makro güvenlik ayarını, parola sorulmasını ve düğmelere basmayı")
    print("GÖREMEZLER. Kurulumdan sonra KURULUM.md'deki elle kontrol")
    print("listesini uygulayın.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
