# Kurulum Talimatnamesi

Bu sürüm için sırayla yapılacaklar.

**Bu sürümde bütün veri `ProjeYonetim.xlsm` dosyasının içindedir.** Ayrı
kayıt dosyası, gelen kutusu, yıl klasörü yoktur.

---

## 1. Ayarları düzenle (üretimden önce)

`kaynak\tasarim.py` — ekranların üstünde yazacak birim adı:

```python
BIRIM_ADI = "XJ Birimi"
```

`kaynak\vba\modAyar.bas` — şifreler ve öneri numarası öneki:

```vba
Public Const SIFRE_PERSONEL As String = "proje"          ' personel ekran şifresi
Public Const SIFRE_YONETIM  As String = "proje-yonetim"  ' ekip ekran şifresi
Public Const SIFRE_DOSYA    As String = "proje-depo"     ' dosya açılış parolası
Public Const ONEK_ONERI_NO  As String = "PRJ"
```

`SIFRE_DOSYA` gizliliğin gerçek sınırıdır — yönetim kitabı bu parolayla
şifrelenir. Diğer ikisi yalnızca kazara girişi önler.

---

## 2. Dosyaları üret

Tek bir geliştirici makinesinde, bir kez. Kullanıcı bilgisayarlarına Python
kurulmaz.

Gerekenler: Windows, masaüstü Excel (365 / 2019+), Python 3.9+.

```
pip install openpyxl pywin32
python kur.py
```

Çıktı:

```
cikti\ProjeOneri.xlsm              → personel açar, şifresiz; öneri gönderir
cikti\ProjeTakip.xlsm              → personel açar, şifresiz; öneri durumunu sorgular
cikti\yonetim\ProjeYonetim.xlsm    → veri deposu, açılış parolalı
```

> ⚠ Üretilen dosyaları **OneDrive ile eşlenmiş bir klasörden çalıştırmayın**
> — denemek için bile. O klasörde değerlendirme kaydedilemez. Denemek için
> `C:\projeoneri\` gibi eşlenmemiş bir klasöre kopyalayın.

---

## 3. Ağ paylaşımına yerleştir

```
\sunucu\paylasim\projeoneri\
├── ProjeOneri.xlsm          ← cikti\ProjeOneri.xlsm
├── ProjeTakip.xlsm          ← cikti\ProjeTakip.xlsm
└── yonetim\
    ├── ProjeYonetim.xlsm    ← cikti\yonetim\ProjeYonetim.xlsm
    └── yedek\               ← elle oluştur, boş bırak
```

`yedek\` klasörünü elle oluşturun; kitap açıldıkça kendi kopyalarını oraya
alır.

`ProjeTakip.xlsm` mutlaka `ProjeOneri.xlsm` ile **aynı klasörde** durmalıdır;
veri deposunu yanındaki `yonetim\` klasöründen bulur.

Yanında zaman zaman `ProjeYonetim.xlsm.kilit` dosyası görünür — yazma
sırasında saniyeden kısa süre var olur, **elle silmeyin.**

---

## 4. İzinleri ver

`Personel` ve `Degerlendirme-Ekibi` yerine kendi grup adlarınızı yazın.
**Sıra önemlidir: önce `yedek\`, sonra `yonetim\`.**

```
set K=\sunucu\paylasim\projeoneri

icacls "%K%\yonetim\yedek" /inheritance:r
icacls "%K%\yonetim\yedek" /grant "ALANADI\Degerlendirme-Ekibi:(OI)(CI)(F)"

icacls "%K%\yonetim" /inheritance:r
icacls "%K%\yonetim" /grant "ALANADI\Degerlendirme-Ekibi:(OI)(CI)(F)"
icacls "%K%\yonetim" /grant "ALANADI\Personel:(OI)(CI)(M)"

icacls "%K%\ProjeOneri.xlsm" /inheritance:r
icacls "%K%\ProjeOneri.xlsm" /grant "ALANADI\Degerlendirme-Ekibi:(F)"
icacls "%K%\ProjeOneri.xlsm" /grant "ALANADI\Personel:(RX)"

icacls "%K%\ProjeTakip.xlsm" /inheritance:r
icacls "%K%\ProjeTakip.xlsm" /grant "ALANADI\Degerlendirme-Ekibi:(F)"
icacls "%K%\ProjeTakip.xlsm" /grant "ALANADI\Personel:(RX)"
```

Personelin `yonetim\` üzerinde **Değiştir** hakkı olması zorunludur — daha
azı gönderimi durdurur. `/t` ile zorla uygulamayın; `yedek\` klasörünün
korumasını bozar.

---

## 5. Excel tarafı (BT ile)

- **Makrolar çalışabilmelidir.** Grup ilkesiyle tümden kapalıysa sistem
  açılmaz.
- **Paylaşımı Güvenilir Konum yapın:** *Dosya → Seçenekler → Güven Merkezi →
  Güven Merkezi Ayarları → Güvenilir Konumlar* → `\sunucu\paylasim\projeoneri`,
  alt klasörler işaretli. "Ağdaki güvenilir konumlara izin ver" kutusu da
  açık olmalı. Yapılmazsa sistem yine çalışır, kullanıcı her açılışta
  *İçeriği Etkinleştir* der.
- **Zorunlu duyarlılık etiketi (MIP) varsa** bu klasör için kaldırılmalıdır;
  yoksa kaydetme penceresi sistemi kilitler.

---

## 6. Kontrol listesi

Kurulumdan sonra bir kez elle:

- [ ] `ProjeOneri.xlsm` ağ yolundan çift tıklayarak açılıyor.
- [ ] *Sisteme Gir* → şifre → form açılıyor.
- [ ] Form doldurulup gönderiliyor, `PRJ-2026-0001` biçiminde numara dönüyor.
- [ ] `ProjeTakip.xlsm` → o numara + sicil no → *Sorgula* → güncel durum ve
      durum geçmişi geliyor.
- [ ] Aynı numara **yanlış sicil no** ile sorgulanınca sonuç gösterilmiyor.
- [ ] Personel hesabıyla `yonetim\ProjeYonetim.xlsm` **parola soruyor.**
- [ ] Personel hesabıyla `yonetim\yedek\` **açılmıyor.**
- [ ] İki kullanıcı aynı anda gönderiyor, ikisi de ayrı numara alıyor.
- [ ] Yönetim kitabı ekip hesabıyla açılıyor, başlıkta **[Salt Okunur]**
      yazıyor. ("Yine de Düzenle" şeridine **basmayın.**)
- [ ] Pano doluyor; *Liste → Önerileri Yenile* gönderimleri gösteriyor.
- [ ] Satıra çift tıkla → durum + karar notu → *Değerlendirmeyi Kaydet*
      çalışıyor.
- [ ] Takip ekranında yeniden sorgulayınca yeni durum geliyor; karar notu
      **görünmüyor.**
- [ ] Ekip kitabı açıkken personel gönderim yapabiliyor.
- [ ] `yonetim\yedek\` altında bugünün kopyası oluştu.

---

## 7. Sonradan değişiklik

Şifre, birim adı, ekran ya da alan değişikliği: kaynağı düzenle, yeniden üret.

**Kurulum doluysa mutlaka veriyi taşıyan biçimi kullanın**, yoksa yeni dosya
boş gelir:

```
python kur.py yonetim --veri "\sunucu\paylasim\projeoneri\yonetim\ProjeYonetim.xlsm"
```

Parola değiştiriyorsanız: önce mevcut kitabın kopyasını alın, `--veri` için o
kopyayı gösterin, `modAyar.bas`'ı kopyayı aldıktan **sonra** düzenleyin.
`SIFRE_DOSYA` değiştiyse `ProjeOneri.xlsm` ve `ProjeTakip.xlsm` de yeniden
üretilip değiştirilmelidir; ikisi de depoyu bu parolayla açar.

---

## 8. Yedekleme

`ProjeYonetim.xlsm` verinin kendisidir; kaybı bütün önerilerin ve
değerlendirme geçmişinin kaybıdır.

- **`yonetim\` klasörünün tamamı kurumsal yedeğe alınmalıdır.** Bu şarttır.
- Sistemin kendi yedeği (`yonetim\yedek\`, günde bir, son yedi kopya) aynı
  diskte durur — kolaylıktır, yeterli değildir.
- `ProjeOneri.xlsm` ve `ProjeTakip.xlsm` yeniden üretilebilir, yedeklenmesi
  gerekmez.
