# Kurulum

Bu belge sistemin bir ağ paylaşımına nasıl kurulacağını anlatır.
Neden böyle tasarlandığı `TASARIM-VE-GEREKCE.md`'de, genel bakış
`OKU-BENI.md`'dedir.

---

## 1. Çalışma kitaplarını üret

Üretim yalnızca **bir kez, bir geliştirici makinesinde** yapılır. Kullanıcı
bilgisayarlarında Python'a da bu adıma da gerek yoktur.

Gerekenler: Windows, masaüstü Excel (Microsoft 365 / 2019+), Python 3.9+.

```
pip install openpyxl pywin32
python kur.py
```

Çıktı:

```
cikti\ProjeOneri.xlsm
cikti\yonetim\ProjeYonetim.xlsm
```

`kur.py` Excel'in **“VBA proje nesne modeline erişime güven”** ayarını geçici
olarak açar ve işi bitince — hata alsa bile — eski değerine döndürür. Bu ayar
kurum ilkesiyle kilitliyse üretim o makinede yapılamaz; başka bir makinede
üretip `.xlsm` dosyalarını kopyalamak yeterlidir.

> **Uyarı:** Üretilen dosyaları OneDrive ile eşlenmiş bir klasörden
> çalıştırmayın. Eşlenmiş klasörlerde Excel dosyanın konumunu disk yolu yerine
> `https://...` adresi olarak bildirir; kod bunu çevirir ama Excel'in Güvenilir
> Konumlar listesi disk yollarına göre çalıştığı için güven ayarı beklendiği
> gibi davranmayabilir. Gerçek kurulum bir UNC paylaşımında olacağı için bu
> durum üretimde oluşmaz.

---

## 2. Klasör yapısını kur

Ağ paylaşımında şu yapıyı oluşturun (örnek: `\\sunucu\paylasim\projeoneri`):

```
projeoneri\
├── ProjeOneri.xlsm          ← cikti\ProjeOneri.xlsm
└── yonetim\
    ├── ProjeYonetim.xlsm    ← cikti\yonetim\ProjeYonetim.xlsm
    ├── oneriler\             ← gönderimlerin TEK ve kalıcı yeri (madde 3)
    └── degerlendirme\
```

Personelin `projeoneri\` altında gördüğü tek şey çalışma kitabı ve `yonetim\`
klasörünün **adıdır**; içine bakamaz.

**Ayrı bir “gelen kutusu” yoktur.** Personel doğrudan `yonetim\oneriler\`
altına yazar; dosya baştan itibaren kalıcı yerindedir ve sonradan hiçbir yere
taşınmaz. Yazabilir ama içini göremez — nasıl olduğu madde 3'te.

**Yıl klasörlerini elle açmayın.** `oneriler\2026\` ve benzerleri ilk
gönderimde kendiliğinden oluşur ve izinleri `oneriler\`den miras alır. Elle
açarsanız izin mirası bozulabilir.

Kitaplar konumlarını **klasör yapısından** bulur: `yonetim` alt klasörü hangi
seviyedeyse kök orasıdır. Dosyaları yeniden adlandırabilirsiniz; yapıyı
bozmayın.

---

## 3. NTFS izinleri — asıl güvenlik sınırı budur

VBA içindeki şifreler gerçek bir sınır değildir (madde 6). Erişimi belirleyen
şey klasör izinleridir.

| Klasör / dosya | Kimler | İzin |
|---|---|---|
| Klasör / dosya | Tüm personel | Değerlendirme ekibi |
|---|---|---|
| `projeoneri\` | Okuma + Çalıştırma | Tam denetim |
| `projeoneri\ProjeOneri.xlsm` | **Salt okunur** | Tam denetim |
| `projeoneri\yonetim\` | **Yalnızca geçiş** (listeleme yok) | Tam denetim |
| `projeoneri\yonetim\oneriler\` | **Bırakma kutusu** — aşağıya bakın | Tam denetim |
| `yonetim\degerlendirme\` | Hiçbir hak | Tam denetim |

### `yonetim\oneriler\` — bırakma kutusu

Bu klasör bir posta kutusu gibi çalışmalıdır: **herkes içine atabilir, kimse
içini göremez.** Personele yalnızca *dosya oluştur* ve *klasör oluştur* hakkı
verilir; *listeleme/okuma* ve *silme* hakkı **verilmez**. Sonuç:

- Kimse başkasının önerisini okuyamaz.
- Kimse klasörde hangi önerilerin olduğunu göremez.
- Kimse bir öneriyi silemez.
- Herkes kendi önerisini bırakabilir.

**Komutların sırası önemlidir: `yonetim\` EN SON kısıtlanır.** Ters sırada
yaparsanız izin komutlarının kendisi çalışamaz hale gelir ve klasörler
sessizce erişilemez kalır (`Personel` ve `Degerlendirme-Ekibi` yerine kendi grup
adlarınızı yazın):

```
set K=\\sunucu\paylasim\projeoneri

REM 1) Önce bırakma kutusu
icacls "%K%\yonetim\oneriler" /inheritance:r
icacls "%K%\yonetim\oneriler" /grant "ALANADI\Degerlendirme-Ekibi:(OI)(CI)(F)"
icacls "%K%\yonetim\oneriler" /grant "ALANADI\Personel:(OI)(CI)(WD,AD,X,RA,REA,WA,WEA,RC)"

REM 2) Sonra yönetim tarafının geri kalanı -- personele hiçbir hak yok
icacls "%K%\yonetim\degerlendirme" /inheritance:r
icacls "%K%\yonetim\degerlendirme" /grant "ALANADI\Degerlendirme-Ekibi:(OI)(CI)(F)"
icacls "%K%\yonetim\ProjeYonetim.xlsm" /inheritance:r
icacls "%K%\yonetim\ProjeYonetim.xlsm" /grant "ALANADI\Degerlendirme-Ekibi:(F)"

REM 3) EN SON: yonetim\ -- personele yalnızca GEÇİŞ, miras bayrağı YOK
icacls "%K%\yonetim" /inheritance:r
icacls "%K%\yonetim" /grant "ALANADI\Degerlendirme-Ekibi:(OI)(CI)(F)"
icacls "%K%\yonetim" /grant "ALANADI\Personel:(X)"
```

Kısaltmalar: `WD` dosya oluştur/veri yaz, `AD` klasör oluştur/veri ekle,
`X` klasörde gezin, `RA`/`REA` öznitelik oku, `WA`/`WEA` öznitelik yaz,
`RC` izinleri oku. **Bilerek verilmeyenler:** `RD` (listele/oku), `DE` (sil),
`DC` (alt öğe sil).

Personele verilen `(X)` hakkında **miras bayrağı yoktur** — yani yalnızca
`yonetim\` klasörünün kendisine uygulanır, kardeş klasörlere sızmaz. Bırakma
kutusundaki izinler ise `(OI)(CI)` ile alt klasörlere ve dosyalara miras
kalır; `/t` ile zorla uygulamayın, mevcut alt klasörlerin izinlerini
boşaltabilir.

> Sistem bu kısıtlar altında **çalışacak şekilde tasarlandı ve sınandı.**
> `python testler\test_izinler.py` klasörleri gerçekten bu izinlerle kilitler,
> gerçek gönderim makrosunu çalıştırır ve şunları tek tek doğrular:
> yönetim klasörü listelenemiyor, değerlendirme notları okunamıyor ve
> üzerlerine yazılamıyor, yönetim kitabı okunamıyor, bırakma kutusu
> listelenemiyor, bırakılan öneri okunamıyor ve silinemiyor — **ama gönderim
> çalışıyor.**
>
> Kayıt yazarken `.tmp` + yeniden adlandırma kullanılmamasının nedeni de
> budur: yeniden adlandırma **silme** yetkisi ister ve bu izin modeliyle
> bağdaşmaz (bkz. `TASARIM-VE-GEREKCE.md`, madde 4).

### `ProjeOneri.xlsm` salt okunur olmalıdır

Aksi halde dosyayı ilk açan kullanıcı kilitler ve ikinci kullanıcı
“kullanımda” uyarısı alır. Salt okunur açılan bir kitap bellekte
düzenlenebilir: form doldurulur, makro ortak klasöre yazar; yalnızca kitabın
kendisi kaydedilemez — zaten istenen budur.

### Bir de gizleme (isteğe bağlı)

`attrib +h "\\sunucu\paylasim\projeoneri\yonetim"` klasörü gözden uzak tutar.
Bu bir güvenlik sınırı **değildir** — asıl koruma yukarıdaki izinlerdir —
ama klasörün merak uyandırmasını önler.

---

## 4. Makro izni — BT ile konuşulması gerekenler

Sistemin tek gerçek gereksinimi makroların çalışabilmesidir. Kuruma sorulacak
iki soru:

1. **Makrolar kullanıcı onayıyla çalışabiliyor mu?** Grup ilkesiyle tümden
   kapatılmışsa bu sürüm hiç açılmaz.
2. **Ağ paylaşımı “Güvenilir Konum” olarak tanımlanabilir mi?**
   - Tanımlanırsa kullanıcı hiçbir uyarı görmez.
   - Tanımlanmazsa sistem yine çalışır; kullanıcı her açılışta
     *İçeriği Etkinleştir* der.

Ek olarak **MOTW (Mark of the Web)**: Ağdan gelen dosyalar bazı yapılandırmalarda
“Korumalı Görünüm”de açılır ve makrolar tümden engellenir. Bunu da Güvenilir
Konum tanımı çözer. Paylaşımın Intranet bölgesinde olması gerekir.

Güvenilir Konum tanımı: *Dosya → Seçenekler → Güven Merkezi → Güven Merkezi
Ayarları → Güvenilir Konumlar → Yeni konum ekle* → `\\sunucu\paylasim\projeoneri`,
“Bu konumun alt klasörlerine de güven” işaretli. (Ağ konumlarına izin vermek
için “Ağdaki güvenilir konumlara izin ver” kutusu da açılmalıdır.)

---

## 5. Birim adını ve şifreleri değiştir

**Birim adı** — ekranların üstündeki lacivert bantta solda yazar.
`kaynak/tasarim.py` başındaki tek satırdır:

```python
BIRIM_ADI = "XJ Birimi"
```

**Ekran şifreleri** — `kaynak/vba/modAyar.bas` başındadır:

```vba
Public Const SIFRE_PERSONEL As String = "proje"
Public Const SIFRE_YONETIM As String = "proje-yonetim"
```

İkisini de değiştirdikten sonra `python kur.py` çalıştırın; kitaplar yeni
değerlerle baştan üretilir.

Öneri numarasının öneki (`PRJ`) `modAyar.bas` içindeki
`ONEK_ONERI_NO` sabitidir. Durum listesini değiştirecekseniz `modModel.bas`
içindeki `Durumlar()` ile `kur.py` içindeki `listeler()` işlevini birlikte
düzenleyin — `python testler\test_uretim.py` ikisinin aynı kaldığını denetler.

---

## 6. Güvenlik — ne gerçek, ne değil

| Katman | Gerçek sınır mı | Not |
|---|---|---|
| Personel şifresi | **Hayır** | VBA içinde düz metin |
| Yönetim şifresi | **Hayır** | VBA içinde düz metin |
| Sayfa/kitap koruması | **Hayır** | Kazara bozmayı önler |
| NTFS klasör izinleri | **Evet** | Asıl erişim denetimi |

Ekran şifrelerinin işlevi “yanlış ekrana yanlışlıkla girmeyi” önlemektir.

---

## 7. Elle uçtan uca kontrol — atlanmamalı

`python testler\tum_testler.py` 145'ten fazla kontrol çalıştırır ve gerçek
Excel'de gerçek makroları kullanır. **Ancak göremediği bir şey vardır:**
makrolar COM üzerinden çağrılır; bu yol Excel'in makro güvenlik ayarını,
“İçeriği Etkinleştir” uyarısını ve düğmelere basmayı hiç görmez.

Kurulumdan sonra aşağıdaki listeyi **bir kez** elle uygulayın:

- [ ] `ProjeOneri.xlsm`'i ağ yolundan çift tıklayarak açın (kendi
      bilgisayarınıza kopyalamadan).
- [ ] Güvenlik uyarısı çıkarsa *İçeriği Etkinleştir*'e basın; çıkmıyorsa
      Güvenilir Konum tanımı çalışıyor demektir.
- [ ] *Sisteme Gir* → şifre → form açılıyor mu?
- [ ] Alanları boş bırakıp *Öneriyi Gönder* → uyarı geliyor mu?
- [ ] Formu doldurup gönderin → öneri numarası görünüyor, form temizleniyor,
      `yonetim\oneriler\<yıl>\` altında dosya oluşuyor mu? (Kontrolü Değerlendirme
      ekibi hesabıyla yapın; personel hesabı klasörü göremez — istenen budur.)
- [ ] Personel hesabıyla `yonetim\` klasörünü açmayı deneyin → **erişim
      engellenmeli.**
- [ ] **İkinci bir kullanıcıyla aynı anda açın** ve ikisi de gönderim yapsın →
      iki ayrı dosya oluşuyor, ikisi de “kullanımda” uyarısı almıyor mu?
- [ ] `yonetim\ProjeYonetim.xlsm`'i açın → *Sisteme Gir* → önce **Pano**
      açılıyor, göstergeler ve grafikler doluyor mu?
- [ ] *Liste* → *Önerileri Yenile* → gönderimler listede mi?
- [ ] Bir satıra çift tıklayın → değerlendirme ekranı doluyor mu?
- [ ] Durum ve karar notu girip *Değerlendirmeyi Kaydet* → geçmişe ekleniyor,
      listede durum değişiyor mu?

Bu liste tamamlanmadan kurulum “bitti” sayılmaz.

---

## 8. Sürdürme

| İş | Nasıl |
|---|---|
| Şifre / başlık / durum listesi değişikliği | `kaynak\vba\*.bas` ya da `kaynak\tasarim.py` düzenle, `python kur.py` |
| Ekran veya alan değişikliği | `kaynak\uret_*.py` düzenle, `python kur.py` |
| Doğrulama | `python testler\tum_testler.py` + madde 7'deki elle liste |
| Tasarım gözden geçirme | `python testler\goruntu_al.py` → ekranların PDF'i |
| **Yedekleme** | **`yonetim\` klasörünün tamamı** (öneriler ve değerlendirmeler orada) |

Yedekleme kuralı önemlidir: çalışma kitapları **üretilebilir** dosyalardır,
veri değildir. Verinin tamamı o iki klasördeki düz metin dosyalarındadır.
`.xlsm` dosyaları her zaman `python kur.py` ile yeniden üretilebilir.

**Ölçek:** `oneriler\` her yenilemede baştan taranır. Yıl alt klasörleri sayesinde
birkaç bin dosyaya kadar sorunsuzdur. On binlere çıkılırsa eski yıl klasörlerini
arşive taşımak yeterlidir; sistem kalan yılları okumaya devam eder.
