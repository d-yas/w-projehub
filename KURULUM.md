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
cikti\ProjeOneri.xlsm                 şifresiz — personel açacak
cikti\yonetim\ProjeYonetim.xlsm       AÇILIŞ PAROLALI — veri deposu
```

Yönetim kitabı `modAyar.bas` içindeki `SIFRE_DOSYA` değeriyle şifrelenir
(madde 5). Elle açmak isterseniz o parola gerekir; sistemin kendisi parolayı
kodun içinden verir, kullanıcıya sormaz.

Testleri de çalıştırmak isterseniz bir bağımlılık daha gerekir — şifreli
kitabın içine bakan yapısal kontroller bunu kullanır:

```
pip install msoffcrypto-tool
python testler\tum_testler.py
```

`kur.py` Excel'in **“VBA proje nesne modeline erişime güven”** ayarını geçici
olarak açar ve işi bitince — hata alsa bile — eski değerine döndürür. Bu ayar
kurum ilkesiyle kilitliyse üretim o makinede yapılamaz; başka bir makinede
üretip `.xlsm` dosyalarını kopyalamak yeterlidir.

> ### ⚠ Üretilen dosyaları OneDrive ile eşlenmiş bir klasörden ÇALIŞTIRMAYIN
>
> Bu bir öneri değil, **kesin bir sınır** ve ölçülerek doğrulandı: OneDrive ile
> eşlenen bir klasördeki çalışma kitabı bir Excel örneğinde açık olduğu sürece
> — salt okunur bile olsa — ikinci bir Excel süreci onu **yazma kipinde
> açamaz.** Excel hata da vermez, sessizce salt okunur açar.
>
> Yönetim kitabı ekipte her zaman açık olduğundan, böyle bir klasörde
> **değerlendirme hiç kaydedilemez.** Gönderim çalışır, değerlendirme
> çalışmaz. Sistem bu durumu tanır: Giriş ekranında uyarı bandı gösterir ve
> kaydetme denemesini beklemeden, sebebini söyleyerek reddeder.
>
> İkinci bir sebep daha var: eşlenmiş klasörlerde Excel dosyanın konumunu disk
> yolu yerine `https://...` adresi olarak bildirir; kod bunu çevirir ama
> Excel'in Güvenilir Konumlar listesi disk yollarına göre çalışır.
>
> **Denemek için bile** çıktıları OneDrive dışında bir klasöre kopyalayın
> (örneğin `C:\projeoneri\`). Gerçek kurulum bir UNC paylaşımında olacağı için
> bu durum üretimde oluşmaz.

---

## 2. Klasör yapısını kur

Ağ paylaşımında şu yapıyı oluşturun (örnek: `\\sunucu\paylasim\projeoneri`):

```
projeoneri\
├── ProjeOneri.xlsm          ← cikti\ProjeOneri.xlsm
└── yonetim\
    ├── ProjeYonetim.xlsm    ← cikti\yonetim\ProjeYonetim.xlsm   (VERİ BURADA)
    └── yedek\               ← boş; kitap ilk açıldığında dolmaya başlar
```

**Bu iki dosya sistemin tamamıdır.** Kayıt dosyası, gelen kutusu, yıl klasörü
yoktur. Bütün öneriler ve bütün değerlendirme geçmişi `ProjeYonetim.xlsm`
dosyasının içindeki iki gizli sayfada durur.

Yanlarında iki yardımcı dosya oluşur ve ikisi de veri taşımaz:

- `yonetim\yedek\ProjeYonetim_YYYYMMDD.xlsm` — günlük yedek kopya.
- `yonetim\ProjeYonetim.xlsm.kilit` — yazma sırasında saniyeden kısa süre var
  olan koordinasyon dosyası. İki kişi aynı anda gönderim yaptığında birinin
  diğerinin satırını ezmesini engeller. Normalde göremezsiniz; bir çökme
  sonrası kalırsa beş dakika içinde kendiliğinden kaldırılır. **Elle
  silmeyin.**

  > Bu kilidin yaşı dosya zaman damgasından hesaplanır, yani **kullanıcı
  > makinelerinin saatleri birbirine yakın olmalıdır.** Alan (domain) ortamında
  > bu zaten sağlanır; alan dışı, saati kaymış bir makine kilidi erken bayat
  > sayabilir.

> ⚠ **`ProjeYonetim.xlsm` verinin kendisidir, üretilebilir bir dosya değildir.**
> Silinirse ya da bozulursa her şey gider. Yedekleme kuralları madde 8'de.

**`yedek\` klasörünü elle oluşturun ve boş bırakın.** Kitap her açıldığında,
günde bir kez, oraya kendi kopyasını alır ve en yeni yedi kopyayı saklar.

Kitaplar konumlarını **klasör yapısından** bulur: `yonetim` alt klasörü hangi
seviyedeyse kök orasıdır. `ProjeOneri.xlsm` yeniden adlandırılabilir; yönetim
kitabının adı `modAyar.bas` içindeki `DOSYA_YONETIM` sabitiyle eşleşmelidir.

---

## 3. NTFS izinleri

**Bu sürümde izin modeli değişti.** Eskiden `yonetim\oneriler\` bir *bırakma
kutusuydu*: personel oraya yazabilir ama içini göremezdi. Veri artık ayrı
dosyalarda değil, yönetim kitabının içinde — ve **Excel bir dosyayı okumadan
yazamaz.** Dolayısıyla personelin o dosya üzerinde okuma hakkı da olmak
zorundadır.

> **Gizliliği sağlayan şey artık klasör izni değil, dosyanın açılış
> parolasıdır.** Personel dosyayı kopyalayabilir ama parolasız açamaz.
> Tehdit modeli sıradan personeldir; VBA'yı açmayı bilen biri parolayı
> okuyabilir (`TASARIM-VE-GEREKCE.md` madde 8).

| Klasör / dosya | Tüm personel | Değerlendirme ekibi |
|---|---|---|
| `projeoneri\` | Okuma + Çalıştırma | Tam denetim |
| `projeoneri\ProjeOneri.xlsm` | **Salt okunur** | Tam denetim |
| `projeoneri\yonetim\` | **Değiştir** (mirasla) | Tam denetim |
| `projeoneri\yonetim\yedek\` | **Hiçbir hak** | Tam denetim |

### Neden "Değiştir", neden klasöre

Excel kaydederken dosyayı yerinde değiştirmez: geçici bir dosya yazıp aslının
yerine koyar. Bunun iki sonucu var:

1. **Personelin klasörde dosya oluşturma ve silme hakkı olmak zorundadır.**
   Daha azı yetmez; gönderim, kitabı kaydetmek demektir. Excel ayrıca `~$` ile
   başlayan bir sahiplik dosyası oluşturup siler ve sistem her yazmada kısa
   ömürlü bir `.kilit` dosyası oluşturup siler.
2. **İzinler dosyaya değil klasöre, mirasla verilmelidir.** Dosyaya doğrudan
   verilen açık haklar kaydetme sırasında kaybolabilir; klasörden miras
   alınanlar kalır. Aksi halde ilk gönderimden sonra dosya erişilemez olur ve
   sistem sessizce durur.

**Komutların sırası önemlidir: `yedek\` ÖNCE, `yonetim\` SONRA.** Ters sırada
yaparsanız izin komutlarının kendisi çalışamaz hale gelir (`Personel` ve
`Degerlendirme-Ekibi` yerine kendi grup adlarınızı yazın):

```
set K=\\sunucu\paylasim\projeoneri

REM 1) Önce yedek klasörü -- personele kapalı
icacls "%K%\yonetim\yedek" /inheritance:r
icacls "%K%\yonetim\yedek" /grant "ALANADI\Degerlendirme-Ekibi:(OI)(CI)(F)"

REM 2) Sonra yonetim\ -- personele Değiştir, dosyalara miras
icacls "%K%\yonetim" /inheritance:r
icacls "%K%\yonetim" /grant "ALANADI\Degerlendirme-Ekibi:(OI)(CI)(F)"
icacls "%K%\yonetim" /grant "ALANADI\Personel:(OI)(CI)(M)"

REM 3) Öneri kitabı personel için salt okunur
icacls "%K%\ProjeOneri.xlsm" /inheritance:r
icacls "%K%\ProjeOneri.xlsm" /grant "ALANADI\Degerlendirme-Ekibi:(F)"
icacls "%K%\ProjeOneri.xlsm" /grant "ALANADI\Personel:(RX)"
```

`M` = Değiştir (oku, yaz, oluştur, sil). `yedek\` klasörünün mirası
kesildiği için `yonetim\` üzerinde personele verilen hak oraya **sızmaz** —
`/t` ile zorla uygulamayın, bu kesmeyi bozar.

> Sistem bu kısıtlar altında **çalışacak şekilde tasarlandı ve sınandı.**
> `python testler\test_izinler.py` klasörleri gerçekten bu izinlerle kilitler,
> gerçek gönderim makrosunu çalıştırır ve şunları tek tek doğrular: gönderim
> çalışıyor, kaydetmeden sonra dosya hâlâ erişilebilir (izinler kaydetmeyi
> atlattı), yedek klasörü listelenemiyor ve üzerine yazılamıyor, yedek
> alınamayınca kullanıcıya hata gösterilmiyor.

### `ProjeOneri.xlsm` salt okunur olmalıdır

Aksi halde dosyayı ilk açan kullanıcı kilitler ve ikinci kullanıcı
"kullanımda" uyarısı alır. Salt okunur açılan bir kitap bellekte
düzenlenebilir: form doldurulur, makro yönetim kitabına yazar; yalnızca
kitabın kendisi kaydedilemez — zaten istenen budur.

### Bir de gizleme (isteğe bağlı)

`attrib +h "\\sunucu\paylasim\projeoneri\yonetim"` klasörü gözden uzak tutar.
Bu bir güvenlik sınırı **değildir** — asıl koruma dosya parolasıdır — ama
klasörün merak uyandırmasını önler.

---

## 4. Makro izni — BT ile konuşulması gerekenler

Sistemin tek gerçek gereksinimi makroların çalışabilmesidir. Kuruma sorulacak
üç soru:

1. **Makrolar kullanıcı onayıyla çalışabiliyor mu?** Grup ilkesiyle tümden
   kapatılmışsa bu sürüm hiç açılmaz.
2. **Ağ paylaşımı "Güvenilir Konum" olarak tanımlanabilir mi?**
   - Tanımlanırsa kullanıcı hiçbir uyarı görmez.
   - Tanımlanmazsa sistem yine çalışır; kullanıcı her açılışta
     *İçeriği Etkinleştir* der.
3. **Kaydetme sırasında zorunlu bir duyarlılık etiketi (MIP) soruluyor mu?**
   Soruluyorsa kaydetme penceresi bastırılamaz ve **sistemin arka planda
   kullandığı görünmez Excel kilitlenir.** Bu durumda yönetim kitabının
   bulunduğu klasör için etiket zorunluluğunun kaldırılması gerekir.

Ek olarak **MOTW (Mark of the Web)**: Ağdan gelen dosyalar bazı yapılandırmalarda
"Korumalı Görünüm"de açılır ve makrolar tümden engellenir. Bunu da Güvenilir
Konum tanımı çözer. Paylaşımın Intranet bölgesinde olması gerekir.

Güvenilir Konum tanımı: *Dosya → Seçenekler → Güven Merkezi → Güven Merkezi
Ayarları → Güvenilir Konumlar → Yeni konum ekle* → `\\sunucu\paylasim\projeoneri`,
"Bu konumun alt klasörlerine de güven" işaretli. (Ağ konumlarına izin vermek
için "Ağdaki güvenilir konumlara izin ver" kutusu da açılmalıdır.)

---

## 5. Birim adını, şifreleri ve depo parolasını değiştir

**Birim adı** — ekranların üstündeki lacivert bantta solda yazar.
`kaynak/tasarim.py` başındaki tek satırdır:

```python
BIRIM_ADI = "XJ Birimi"
```

**Ekran şifreleri ve depo parolası** — `kaynak/vba/modAyar.bas` başındadır:

```vba
Public Const SIFRE_PERSONEL As String = "proje"
Public Const SIFRE_YONETIM As String = "proje-yonetim"
Public Const SIFRE_DOSYA As String = "proje-depo"
```

`SIFRE_DOSYA` yönetim kitabının **açılış parolasıdır** ve gizliliğin gerçek
sınırıdır (madde 6). `kur.py` bu değeri doğrudan buradan okur; başka bir yerde
kopyası yoktur.

Değiştirdikten sonra `python kur.py` çalıştırın. **Kurulum zaten doluysa**
veriyi taşıyan biçimi kullanın, yoksa yeni dosya boş gelir:

```
python kur.py yonetim --veri "\\sunucu\paylasim\projeoneri\yonetim\ProjeYonetim.xlsm"
```

Eski kitap **eski parolayla** okunur ve içeriği yeni parolayla kaydedilmiş
kitaba taşınır. Bu yüzden parolayı değiştirirken önce eski kitabın bir
kopyasını alın; `--veri` için o kopyayı gösterin ve `modAyar.bas`'ı kopyayı
aldıktan **sonra** düzenleyin.

Öneri numarasının öneki (`PRJ`) `modAyar.bas` içindeki `ONEK_ONERI_NO`
sabitidir. Durum listesini değiştirecekseniz `modModel.bas` içindeki
`Durumlar()` ile `kur.py` içindeki `listeler()` işlevini birlikte düzenleyin —
`python testler\test_uretim.py` ikisinin aynı kaldığını denetler.

---

## 6. Güvenlik — ne gerçek, ne değil

**Tehdit modeli sıradan personeldir.** Amaç, bir çalışanın başkalarının
önerilerini kazara ya da merakla görmesini engellemektir.

| Katman | Gerçek sınır mı | Not |
|---|---|---|
| Personel / yönetim ekran şifresi | **Hayır** | VBA içinde düz metin |
| Sayfa ve kitap koruması | **Hayır** | Kazara bozmayı önler |
| **Yönetim kitabının açılış parolası** | **Evet** (sıradan personele karşı) | Dosya şifreli; parolasız açılmaz |
| NTFS izinleri | **Kısmen** | Yalnızca `yedek\` klasörünü korur |

Personel `yonetim\` klasörünü ve dosya adlarını görebilir, dosyanın içini
göremez. Kopyalayabilir ama açamaz. Teknik olarak silebilir — bunun karşılığı
yedeklerdir.

---

## 7. Elle uçtan uca kontrol — atlanmamalı

`python testler\tum_testler.py` 250'den fazla kontrol çalıştırır ve gerçek
Excel'de gerçek makroları kullanır. **Ancak göremediği şeyler vardır:**
makrolar COM üzerinden çağrılır; bu yol Excel'in makro güvenlik ayarını,
"İçeriği Etkinleştir" uyarısını, parola sorulmasını ve düğmelere basmayı hiç
görmez.

Kurulumdan sonra aşağıdaki listeyi **bir kez** elle uygulayın:

- [ ] `ProjeOneri.xlsm`'i ağ yolundan çift tıklayarak açın (kendi
      bilgisayarınıza kopyalamadan).
- [ ] Güvenlik uyarısı çıkarsa *İçeriği Etkinleştir*'e basın; çıkmıyorsa
      Güvenilir Konum tanımı çalışıyor demektir.
- [ ] *Sisteme Gir* → şifre → form açılıyor mu?
- [ ] Alanları boş bırakıp *Öneriyi Gönder* → uyarı geliyor mu?
- [ ] Formu doldurup gönderin → birkaç saniye sonra `PRJ-2026-0001` biçiminde
      bir numara görünüyor ve form temizleniyor mu?
- [ ] Personel hesabıyla `yonetim\ProjeYonetim.xlsm`'i açmayı deneyin →
      **parola sorulmalı**, parola olmadan açılmamalı.
- [ ] Personel hesabıyla `yonetim\yedek\` klasörünü açmayı deneyin →
      **erişim engellenmeli.**
- [ ] **İkinci bir kullanıcıyla aynı anda açın** ve ikisi de gönderim yapsın →
      iki ayrı numara alıyor, hiçbiri hata almıyor mu?
- [ ] `yonetim\ProjeYonetim.xlsm`'i ekip hesabıyla açın → parola → başlık
      çubuğunda **[Salt Okunur]** yazıyor mu? (Yazmıyorsa kilit ekipte kalmış
      demektir; Giriş ekranında uyarı bandı görünür.)
- [ ] Excel bir "Yine de Düzenle" şeridi gösterirse **basmayın**: kilidi geri
      alır ve personel öneri gönderemez.
- [ ] *Sisteme Gir* → önce **Pano** açılıyor, göstergeler ve grafikler doluyor mu?
- [ ] *Liste* → *Önerileri Yenile* → gönderimler listede mi?
- [ ] Bir satıra çift tıklayın → değerlendirme ekranı doluyor mu?
- [ ] Durum ve karar notu girip *Değerlendirmeyi Kaydet* → geçmişe ekleniyor,
      listede durum değişiyor mu?
- [ ] **Ekip kitabı açıkken** personel bir öneri daha göndersin → çalışıyor mu?
- [ ] Yönetim kitabında Ctrl+S deneyin → kaydetme engellendi mi?
- [ ] `yonetim\yedek\` altında bugünün tarihli kopyası oluştu mu?

Bu liste tamamlanmadan kurulum "bitti" sayılmaz.

---

## 8. Sürdürme ve yedekleme

| İş | Nasıl |
|---|---|
| Şifre / parola / başlık / durum listesi değişikliği | `kaynak\vba\*.bas` ya da `kaynak\tasarim.py` düzenle, `python kur.py` |
| Ekran veya alan değişikliği | `kaynak\uret_*.py` düzenle, `python kur.py` |
| **Dolu bir kurulumu güncelleme** | **`python kur.py yonetim --veri <mevcut ProjeYonetim.xlsm>`** |
| Doğrulama | `python testler\tum_testler.py` + madde 7'deki elle liste |
| Tasarım gözden geçirme | `python testler\goruntu_al.py` → ekranların PDF'i |

### Yedekleme — bu sürümde kural değişti

**`ProjeYonetim.xlsm` verinin kendisidir.** Eski sürümde çalışma kitapları
üretilebilir dosyalardı ve veri klasördeki metin dosyalarındaydı; artık öyle
değil. O dosyanın kaybı bütün önerilerin ve bütün değerlendirme geçmişinin
kaybıdır.

- **Kurumsal yedek şarttır.** `yonetim\` klasörünün tamamı düzenli olarak
  yedeklenmelidir.
- **Sistemin kendi yedeği bir kolaylıktır, yeterli değildir.** Kitap her
  açıldığında günde bir kez `yonetim\yedek\ProjeYonetim_YYYYMMDD.xlsm` kopyası
  alınır ve en yeni yedi kopya saklanır. Aynı diskte durur; disk giderse o da
  gider.
- `ProjeOneri.xlsm` hâlâ üretilebilir bir dosyadır, yedeklenmesi gerekmez.

### Ölçek

Birkaç bin satıra kadar sorun beklenmez. Her gönderim dosyanın tamamını
yeniden yazar (ölçülen süre ~3 saniye) ve yazmalar sıraya girer; dosya
büyüdükçe bu süre uzar. On
binlere çıkılırsa eski yılların satırları bir arşiv kitabına taşınmalıdır.
