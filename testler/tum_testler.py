# -*- coding: utf-8 -*-
r"""Butun testleri sirayla calistirir.

    python testler\tum_testler.py

Sira ucuzdan pahaliya dogrudur: yapisal kontroller Excel acmadan saniyeler
icinde biter, uctan uca ve eszamanlilik testleri gercek Excel ornekleri
baslatir. Bir asama basarisiz olursa sonrakiler yine de calisir; boylece tek
kosuda butun sorunlar gorulur.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import test_eszamanlilik
import test_izinler
import test_uctan_uca
import test_uretim


def main():
    asamalar = [
        ("Üretim", test_uretim.calistir),
        ("Uçtan uca", test_uctan_uca.calistir),
        ("Eşzamanlılık", test_eszamanlilik.calistir),
        ("İzinler", test_izinler.calistir),
    ]

    kalan = []
    for ad, calistir in asamalar:
        print(f"\n### {ad}")
        if calistir() != 0:
            kalan.append(ad)

    print()
    if kalan:
        print(f"BAŞARISIZ aşamalar: {', '.join(kalan)}")
        return 1
    print("Tüm aşamalar geçti.")
    print()
    print("Hatırlatma: bu testler makroları COM üzerinden çağırır; Excel'in")
    print("makro güvenlik ayarını ve düğmelere basmayı GÖREMEZLER.")
    print("Kurulumdan sonra KURULUM.md'deki elle kontrol listesini uygulayın.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
