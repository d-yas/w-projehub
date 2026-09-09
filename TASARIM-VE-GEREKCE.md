# Proje Öneri Sistemi — Sürüm 2 (İki Excel Dosyası)
## Tasarım, gerekçe ve kısıtlar

Bu belge Sürüm 2'nin **neden böyle tasarlandığını** anlatır. Nasıl kurulacağı
`KURULUM.md`'de, sistemin genel tanıtımı `OKU-BENI.md`'dedir.

---

## 1. Bir bakışta

|  |  |
|---|---|
| **Amaç** | Personelin iyileştirme önerisi göndermesi, Değerlendirme ekibinin bunları PDCA döngüsüyle takip etmesi |
| **Yaklaşım** | Sunucu yok, veritabanı yok, kayıt dosyası yok. **Yalnızca iki Excel dosyası** |
| **Kullanıcıda kurulum** | Yok. Herkeste zaten olan Excel yeterli |
| **Tek gerçek gereksinim** | Makroların çalışmasına izin verilmesi |
| **Durum** | Tamamlandı; 250'den fazla otomatik kontrolle, gerçek Excel'de doğrulandı |
| **Karar bekleyen** | Banka BT'sinin makro ve güvenilir konum politikası (madde 10) |

**Neden bu sürüm var:** Sürüm 1 sürekli açık bir bilgisayarda Flask sunucusu
çalıştırır. Böyle bir makine tahsis edilemezse ya da BT ağda uygulama sunucusu
istemezse tüm sistem çöker. Sürüm 2 aynı işi **hiçbir sunucu olmadan** yapar;
iki sürüm birbirinden bağımsızdır, hangisi uygunsa o seçilir.

**Sürüm 2 içinde yapılan büyük değişiklik.** Sürüm 2 önce her gönderimi ve her
değerlendirmeyi ortak klasöre ayrı bir metin dosyası olarak yazıyordu. Bu
katman tümüyle kaldırıldı: sistemde artık **hiçbir kayıt dosyası yok**, veri
Yönetim kitabının içindeki iki gizli sayfada duruyor. Gerekçesi, bedelleri ve
yerini alan eşzamanlılık kuralı 3. ve 4. maddelerde.

---

## 2. Hedeflenen

1. Personel öneri gönderirken hiçbir şey kurmasın, hiçbir yere kaydolmasın
2. İki kişi aynı anda gönderim yaptığında hiçbir öneri kaybolmasın
3. Değerlendirme ekibi önerileri tek ekrandan görsün, değerlendirsin, önceliklendirsin
4. "Kim, ne zaman, neyi değiştirdi" izi kaybolmasın — bankacılıkta aranan izlenebilirlik
5. Sistem bozulsa bile veri kaybolmasın

---

## 3. Kısıtlar ve bunların tasarımı nasıl belirlediği

Tasarımın tamamı dört kısıtın sonucudur. Her karar bir kısıta cevaptır.

### Kısıt 1 — Sunucu yok

Sürekli açık bir makine ve WSGI sunucusu varsayılamıyor.

> **Sonuç:** Bütün mantık istemci tarafında, kullanıcının kendi Excel'inde çalışır.
> Ağ paylaşımı yalnızca iki dosyanın durduğu yerdir; üzerinde çalışan bir
> program yoktur.

### Kısıt 2 — `file://` ile açılan HTML sayfası diske yazamaz

İlk akla gelen çözüm Sürüm 1'in HTML sayfalarını olduğu gibi ortak klasöre
koymaktı. Çalışmaz: tarayıcılar `file://` protokolüyle açılan bir sayfanın
diske yazmasını güvenlik gereği engeller. Sunucu olmadan HTML'in yazma yolu yok.

> **Sonuç:** Excel + VBA seçildi. Excel yazma yetkisine sahiptir, herkeste
> kuruludur ve banka personeli zaten Excel biliyor.

### Kısıt 3 — Aynı dosyaya iki kişi aynı anda yazamaz

Bir kullanıcı Excel dosyasını yazma kipinde açtığında dosyayı kilitler.

Bu kısıtın **iki farklı cevabı** denendi ve ikincisi seçildi.

*Önceki cevap — kimse ortak dosyaya yazmaz.* Her gönderim `yonetim\oneriler\`
klasörüne kendi metin dosyasını bırakıyordu; çakışma önlenmiş değil, yapısal
olarak imkânsızdı. Bedeli, kullanıcının hiç görmediği bir dosya ve klasör
katmanıydı: yıl klasörleri, kayıt biçimi, yarım dosya koruması, UTF-8
kodlaması, klasör tarama. Bu katmanın bakımı sistemin geri kalanından fazlaydı
ve kullanıcı bunu istemedi.

> **Bugünkü kural:** *Herkes aynı dosyaya, kısa ve dışlayıcı bir işlemle yazar.*
>
> ```
> kilidi al → aç (yazma kipi) → satır ekle → kaydet → kapat → kilidi bırak
>                                                          ≈ 3 saniye
> ```
>
> Çakışmada **geri çekilinir, rastgele 300–900 ms beklenir, yeniden denenir**;
> sıra beklemek için bütçe 60, kilit alındıktan sonra yazmak için 20
> saniyedir. Bütçe dolarsa kullanıcı anlaşılır bir hata görür ve
> **yazdıkları formda durur.**

Bu güvence **üç katmandan** oluşur ve üçü de gereklidir. İkisi ölçüm sırasında
yetersiz çıktığı için eklendi; hangisinin neyi yakaladığı önemlidir.

**1. Kilit dosyası — gerçek karşılıklı dışlama.** Excel'in kendi dosya kilidi
tek başına **yetmez**, çünkü Excel kaydederken dosyayı yerinde değiştirmez:
geçici bir dosya yazıp aslının *yerine koyar*. O yer değiştirme anında eski
dosyanın kilidi bırakılır, yenisininki henüz alınmamıştır. Pencere milisaniyeler
sürer ama dört taraf aynı anda yazıyorsa yakalanır:

```
A: aç → en büyük numara = N → satır N+1 → KAYDET ─┐
B:                        aç (pencere!) ──────────┴→ eski içeriği görür
B: en büyük numara = N → satır N+1 → KAYDET  → A'nın satırını EZER
```

Sonuç: bir satır kaybolur ve **aynı numara iki öneriye verilir.** Eşzamanlılık
testi bunu tam olarak böyle yakaladı: dokuz gönderim, sekiz numara, yirmi sekiz
satır. Bu yüzden yazmaya başlamadan önce `yonetim\` altında
`ProjeYonetim.xlsm.kilit` adlı bir dosya **"varsa oluşturma" kipinde** yaratılır
(`FileSystemObject.CreateTextFile(..., Overwrite:=False)`); bu işlem ağ
paylaşımlarında da atomiktir. Sıfır baytlık, geçici bir **koordinasyon**
dosyasıdır — veri taşımaz ve işlem biter bitmez silinir. Sahibi çökerse iki
dakika sonra bayat sayılıp kaldırılır.

**2. Açılan kitabın `ReadOnly` özelliği.** Dosya başkası tarafından yazma
kipinde açıkken `Workbooks.Open(..., Notify:=False)` çağrısı, uyarılar
kapalıyken **hata vermez: dosyayı sessizce salt okunur açar.** Yalnızca hataya
bakan bir kod, yazdığını sanıp hiçbir şey yazmamış olurdu.

**3. `Save`'in sonucu.** Kaydetme hatası yutulursa satır diske inmediği halde
çağırana başarı döner; üstelik sonraki yazıcı aynı "en büyük + 1" değerini
hesaplar. Bu yüzden `Save` asla varsayılmaz, sonucu denetlenir ve başarısızlıkta
**bütün işlem baştan denenir.**

**Ekip kitabı kilidi tutmaz.** Değerlendirme ekibi Yönetim kitabını gün boyu
açık tutar. Hiçbir şey yapılmasaydı bu, personelin gün boyu öneri
gönderememesi demekti. `Workbook_Open` ilk iş olarak `ChangeFileAccess` ile
kitabı salt okunur kipe alır ve kilidi bırakır; ekran çalışmaya devam eder,
yalnızca kaydedilemez.

### Kısıt 4 — VBA içindeki parola gerçek bir güvenlik sınırı değildir

Parolalar kodun içinde durur. VBA proje parolaları ücretsiz araçlarla
kırılabilir.

> **Sonuç:** Ekran şifreleri yalnızca "yanlış ekrana yanlışlıkla girmeyi"
> engeller. Yönetim kitabının **açılış parolası** ise gerçek bir engeldir ama
> yalnızca **sıradan personele karşı**: dosyayı kopyalayan biri onu parolasız
> açamaz, VBA'yı açmayı bilen biri parolayı okur. Tehdit modeli bilinçli
> olarak budur (madde 8).

---

## 4. Nasıl çalışır

```
                 PERSONEL                             KAİZEN EKİBİ
            ProjeOneri.xlsm                    yonetim\ProjeYonetim.xlsm
         (NTFS salt okunur, şifresiz)          (açılış parolalı — VERİ BURADA)
                    │                                       │
         şifre → form → Gönder                  parola → Pano → Liste
                    │                                       │
                    │  kilidi alır, gizli bir Excel         │  Workbook_Open:
                    │  örneği dosyayı açar, bir satır       │   1) günlük yedek
                    ▼  ekler, kaydeder, kapatır,            │   2) SALT OKUNURA GEÇ
        ┌───────────────────────────┐  kilidi bırakır.      │      (kilidi bırakır)
        │  ProjeYonetim.xlsm        │ ◄─────────────────────┘
        │   Oneriler  (değişmez)    │   Yenile  → diskten çeker
        │   Olaylar   (ekle-only)   │   Kaydet  → yeni olay satırı
        │   Veri / PanoVeri / ekran │
        └───────────────────────────┘
                    │
                    ▼
        yonetim\yedek\ProjeYonetim_YYYYMMDD.xlsm   (günde bir, son 7)
```

**Sistem iki dosyadan ibarettir.** Kayıt dosyası, gelen kutusu, yıl klasörü
yok. Yanlarında üretilen iki yardımcı dosya vardır ve ikisi de veri taşımaz:
günlük **yedek kopya** (`yedek\` altında) ve yazma sırasında saniyeden kısa
süre var olan **kilit dosyası** (`ProjeYonetim.xlsm.kilit`, madde 3).

**Veri Yönetim kitabının içindedir**, iki *çok gizli* sayfada:

| Sayfa | Ne tutar | Kural |
|---|---|---|
| `Oneriler` | Satır başına bir gönderim | **Değişmez.** Yazıldıktan sonra hiç güncellenmez |
| `Olaylar` | Satır başına bir değerlendirme | **Yalnızca eklenir.** Hiçbir satır silinmez ya da değiştirilmez |

Sütunları `Oneriler` için: `sema`, `oneri_no`, `tarih`, `ad_soyad`,
`sicil_no`, `mevcut_durum`, `oneri_basligi`, `cozum_onerisi`,
`beklenen_fayda`, `gonderen_bilgisayar`, `gonderen_kullanici`. `Olaylar`
için: `sema`, `oneri_no`, `olay_tarihi`, `yeni_durum`, `karar_notu`,
`degerlendiren_kullanici`, `degerlendiren_bilgisayar`.

**Gönderim kaydında `durum` alanı yoktur.** Durum yalnızca olaylardan
türetilir; aynı bilgi iki yerde tutulmaz.

**Bütün sütunlar metin biçimlidir.** Aksi halde Excel `10045` sicil numarasını
sayıya, `2026-09-08T15:34:27` tarihini tarihe çevirir; baştaki sıfırlar ve
saniye bilgisi sessizce kaybolur. `=` ile başlayan bir metnin formül sanılması
da böyle engellenir.

**Çok satırlı alanlar hücre içi satır sonuyla durur.** Eski tasarımda satır
sonları `<|>` gibi bir belirtece çevriliyordu, çünkü kayıt tek satırlık metin
olmak zorundaydı. Hücrede böyle bir kısıt yok; metin olduğu gibi saklanır.

**Yazma neden ayrı, gizli bir Excel örneğinden yapılır.** Üç sebep: (a) Yönetim
kitabı ekibin ekranında zaten açık olabilir ve aynı dosya aynı Excel örneğinde
ikinci kez açılamaz; (b) kullanıcının kendi Excel'inin ayarlarına hiç
dokunulmaz, bir hata olsa bile onun Excel'i bozuk ayarlarla kalmaz; (c) depo
kitabının `Workbook_Open` kodu `AutomationSecurity` ile tümden kapatılır.
Bedeli örnek başına yaklaşık bir saniyedir.

**Öneri numarası `PRJ-2026-0001` biçimindedir.** Sıralı sayaç eskiden
*imkânsızdı*: ortak bir sayaç dosyası gerektiriyordu ve "kimse ortak dosyaya
yazmaz" kuralı buna izin vermiyordu. O yüzden numara rastgeleydi (`PRJ-26A7K`).
Artık yazma zaten dışlayıcı bir kilit altında yapılıyor; sayacı okuyup bir
artırmak yarışsızdır. Numara hem okunur hem de kaçıncı önerinin geldiğini
doğrudan gösterir.

**Ekle-only güvencesi artık satır sırasından gelir.** Eski tasarımda olayların
sırası dosya adındaki 10 ms çözünürlüklü zaman damgasından okunuyordu; aynı
saniyede yazılan iki olayın sırası belirsiz kalabiliyordu. Satırlar yalnızca
sona eklendiği ve ekleme kilit altında yapıldığı için sıra artık kesindir.

**Yönetim kitabı ASLA kaydedilmez.** Ekip kitabı salt okunur açar; ekranda
yapılan her şey yalnızca bellektedir. Bir kaydetme denemesi iki türlü zarar
verebilirdi: ya reddedilip kullanıcıyı *Farklı Kaydet*'e yönlendirir ve depo
ikiye bölünürdü, ya da bellekteki **eski** kopya diskin üzerine yazılır ve o
arada personelin gönderdiği bütün öneriler silinirdi. `Workbook_BeforeSave`
kaydetmeyi iptal eder ve nedenini söyler.

**Ekran her zaman kitabın kendi kopyasından okur.** *Önerileri Yenile* önce
diskteki `Oneriler` ve `Olaylar` satırlarını bu kitaba çeker, sonra tabloyu ve
panoyu baştan kurar. Gizli `Veri` sayfası bu türetmenin sonucudur — bir
önbellektir, kaynak değildir.

**Günlük yedek tek gerçek sigortadır.** Veri tek bir dosyada durduğu için o
dosyanın silinmesi ya da bozulması her şeyin gitmesi demektir. Kitap her
açıldığında, günde bir kez, `yonetim\yedek\` altına bayt bayt aynı bir kopya
alınır ve en yeni yedisi saklanır. Kopya `SaveCopyAs` ile değil **disk
kopyasıyla** alınır: `SaveCopyAs` bellekteki durumu yazar, disk kopyası ise
dosyanın aynısıdır ve parolası da yerindedir.

**Yeniden üretmek veriyi silebilir — bunun bir yolu vardır.** Yönetim kitabı
artık veri deposu olduğu için `python kur.py` onu boş olarak yeniden üretir.
Dolu bir kurulumu güncellerken `python kur.py yonetim --veri <mevcut kitap>`
kullanılır; eski kitaptaki bütün satırlar yenisine taşınır.

## 5. Değerlendirme modeli

Sistem yalnızca bir form değil; bir iyileştirme yöntemi uygular.

**Durum akışı (PDCA):**

| Durum | PDCA | Anlamı |
|---|---|---|
| Yeni | — | Henüz incelenmedi |
| Değerlendirmede | Planla | Değerlendirme ekibi inceliyor |
| Planlandı | Planla | Kabul edildi, uygulama planlanıyor |
| Pilot Uygulamada | Uygula | Sınırlı alanda deneniyor |
| Ölçümleniyor | Kontrol | Pilot sonuçları ölçülüyor |
| Standartlaştırıldı | Önlem | Yaygınlaştırıldı, süreç haline geldi |
| Beklemede / Reddedildi | — | Şartlar oluşmadı / gerekçesiyle kapatıldı |

**Kapsam dışı bırakılanlar.** İlk tasarımda formda iki alan daha vardı: gönderen
birimi ve israf (muda) türü. İkisi de kaldırıldı. Birim, sicil numarasından
zaten çıkarılabilecek bir bilgiydi; israf türü ise gönderen personelden
sınıflandırma yapmasını istiyordu ve bu, formu doldurmayı gereksiz
zorlaştırıyordu. Öneri modelinin ağırlığı PDCA durum akışında kaldı.

**Puanlama ve tasarruf ölçümü de kaldırıldı.** Bir sürüm boyunca etki/efor
puanı (1–5), bundan türeyen dört kutuluk öncelik matrisi ve yıllık saat / TL
tasarruf alanları vardı. Üçü de değerlendiren kişiden, elinde ölçüm olmadan
sayı üretmesini istiyordu; girilmeyince pano boş, girilince güvenilmez
oluyordu. Değerlendirme artık iki alandır: **yeni durum** ve **karar notu /
yorum**. Öncelik, durum akışının kendisinden okunur.

**Göstergeler** panoda otomatik hesaplanır: toplam öneri, değerlendirme
bekleyen, uygulamaya geçmiş ve bu ay gelen sayıları; ayrıca durum dağılımı ve
son 12 ayın gönderim trendi grafikleri. Hepsi durumdan türer — elle girilen
hiçbir rakama dayanmaz.

**Pano açılış ekranıdır.** Yönetim kitabında şifre girildiğinde ilk gelen ekran
Pano'dur; kitaptaki ilk sayfa da odur. Gerekçe: ekibin günlük ilk sorusu "ne
durumdayız"dır, "hangi öneri" değil. Ayrıntıya inmek bir düğme uzaktadır.
Öneri tablosunun adı **Liste**'dir (bir sürüm boyunca "Konsol" idi); ekranın
yaptığı iş tam olarak budur ve ekip de onu böyle adlandırıyordu.

**Rapor ekranı kaldırıldı.** Bir sürüm boyunca ayrı bir ekran, dönem süzgeciyle
`yonetim\rapor\` altına düz bir `.xlsx` üretiyordu (daha öncesinde bu dosya AES
parolalıydı; parola üretim sırasında sorulup hiçbir yere kaydedilmiyordu ve
kaybolduğunda rapor kurtarılamıyordu — o da kaldırılmıştı). Panonun göstergeleri
ve listenin süzülebilir tablosu aynı soruları anında yanıtladığı için ekran
kullanılmadı; ayrı bir dosya üretmek, onu ayrı bir klasörde izinlerle korumak ve
klasörün dışına çıktığı anda korumasız kalmasını göze almak karşılıksız bir
maliyetti. **Bedeli açıktır:** dönem süzgeçli hazır bir çıktı yok; yönetime bir
şey gidecekse liste tablosu doğrudan kopyalanır.

---

## 6. Görsel dil

Ekranlar bir hesap tablosu gibi değil, bir uygulama gibi okunmalı. Bunu dört
karar sağlıyor.

**Zemin açık, içerik beyaz.** Sayfa tek renk (`F5F7FA`) boyanır; okunacak her
şey — kart, tablo gövdesi, giriş alanı — beyaz bir yüzeye oturur. Katman farkı,
kutu çizmeden hiyerarşi kurar. Marka laciverti (`002D62`) değişmedi ama artık
yalnızca üç yerde kullanılıyor: başlık bandı, birincil düğme ve KPI rakamı.
Tablo başlıkları dolu lacivert bir banttan açık zemine geçti; koyu bant,
tablonun kendisinden daha çok dikkat çekiyordu. Çizgiler üç basamağa ayrıldı:
`CIZGI_INCE` (tablo satırları), `CIZGI_GRI` (kart kenarı), `ALAN_CIZGI` (giriş
alanı). Tablolardan dikey çizgiler tamamen kaldırıldı — "hücre ızgarası"
hissini en çok üreten şey onlardı.

**Yazılabilir olan beyazdır.** Değerlendirme ekranında salt okunur alanların
kutusu yoktur; zeminde dururlar ve altlarında ince bir çizgi vardır. Beyaz
yüzey yalnızca yazılabilir alanlara ayrılmıştır. "Buraya yazamazsın" bilgisi
bir kilit uyarısıyla değil, görünümle verilir.

**Pano'da gerçek şekiller.** Excel hücrelere yuvarlak köşe ve yumuşak gölge
veremez; şekiller verebilir. Bedeli şudur: **şekiller her zaman hücrelerin
üstünde çizilir**, yani kartın metni de şekle taşınmak zorundadır. Bu yüzden
Pano'da iki katman var. Altta hücreler: göstergelerin **tek doğruluk kaynağı**
orasıdır, `modPano` oraya yazar ve testler oradan okur. Üstte, aynı sayıyı
gösteren şekiller: `modPano.KartYaz` her yenilemede şeklin metnini hücredeki
değerle eşitler. Şekil metnini hücreye *formülle* bağlamak da mümkündü ama
denendi ve elendi: bağlı metin hücrenin sayı biçimini yok sayıp `1234` yerine
`1234,0` yazıyor. Şekil katmanı bir Excel sürümünde kurulamazsa `try/except`
onu atlar; gösterge yine doğru hesaplanır, yalnızca kart yüzeyi düz görünür.
Grafiklerin arkasındaki paneller de aynı şekilde çizilir ve **grafiklerden önce**
eklenmek zorundadır: şekiller ile grafikler aynı çizim katmanındadır ve ekleme
sırasına göre üst üste binerler.

**Excel'in kendi arayüzü olduğu gibi kalır.** Bir sürüm boyunca oturum
açıldığında formül çubuğu, sayfa sekmeleri ve satır/sütun başlıkları
gizleniyordu ("kiosk modu"); ekran gerçekten bir uygulamaya benziyordu ama
günlük kullanımda engel oldu — bu ayarların bir kısmı **uygulama düzeyindedir**
ve aynı Excel'de açık olan diğer kitapları da etkiler. Kaldırıldı. Ekranın
uygulama gibi görünmesi artık yalnızca sayfanın kendi ayarlarına dayanıyor:
kılavuz çizgileri kapalı, tek renk zemin, beyaz yüzeyler. Bunlar dosyanın
içinde durur, kullanıcının Excel'ine dokunmaz.

> **VBA tuzağı — pahalıya mal oldu, tekrarlanmasın.** Modül düzeyi değişkenler
> VBA'da yalnızca *bildirim bölümünde*, yani ilk `Sub`/`Function`'dan önce
> durabilir. Yordamların arasına konan bir `Private x As Boolean` modülü
> **derletmez**; `python kur.py`'nin düğme doğrulaması bunu görmez, çünkü VBA
> yordamları isteğe bağlı derler. Hata ancak o yordam ilk kez çağrıldığında
> ortaya çıkar — görünmeyen bir Excel'de ise hiç görünmez: makro sessizce geri
> dönmez. `test_uretim.py` artık her `.bas` dosyasında bu yerleşimi denetler.

---

## 7. Teknik yapı

| Modül | Sorumluluk |
|---|---|
| `modAyar.bas` | Dosya yolları, ekran şifreleri, **depo parolası**, şema sürümü |
| `modTasarim.bas` | Renk ve tipografi sabitleri, durum rozetleri |
| `modDosyaIO.bas` | Yol işlemleri ve OneDrive yol çevirisi (yalnızca bu kaldı) |
| **`modDepo.bas`** | **Veri deposu: kilit, yeniden deneme, gizli Excel örneği, sıralı numara, günlük yedek** |
| `modModel.bas` | Değerlendirme modeli: durumlar, PDCA, durum anlamları |
| `modUI.bas` | Ekran yönetimi, şifre kapısı, koruma, mesajlar, sessiz mod |
| `modGonderim.bas` | Personel tarafı: form doğrulama, gönderim |
| `modKonsolide.bas` | Depo sayfalarını okuyup olayları tekrar oynatarak tabloyu kurma |
| `modDegerlendirme.bas` | Değerlendirme ekranı, ekle-only olay satırı, geçmiş |
| `modPano.bas` | Göstergeler ve grafik kaynak verisi |

`modDepo` iki kitapta da bulunur: gönderim tarafı da aynı depoya yazar.
Eşzamanlılığa dair bütün risk tek bir modülde toplanmıştır.

**`modDosyaIO` neden küçüldü:** bir zamanlar sistemin veri katmanıydı —
UTF-8 akışları, atomik yazma, kayıt biçimi, klasör tarama, zaman damgası.
Kayıt dosyaları kalkınca hepsi gereksizleşti ve silindi; geriye yalnızca
OneDrive yol çevirisi ve birkaç yol yardımcısı kaldı.

Sayfa ve `ThisWorkbook` kod-arkası en az düzeyde tutulur (yalnızca olay
yönlendirme); tüm mantık modüllerde kalır ki `.bas` dosyaları tek kaynak olsun.

Çalışma kitapları elle hazırlanmaz. **`kur.py`** iki `.xlsm` dosyasını sıfırdan
üretir. İki aşamalıdır:

1. **openpyxl** görsel iskeleti kurar: sayfalar, hücre birleşimleri, ölçüler,
   renkler, koşullu biçimlendirme, doğrulama listeleri, adlandırılmış aralıklar.
2. **pywin32 (COM)** gerçek Excel üzerinden geri kalanı tamamlar: VBA modülleri,
   yuvarlatılmış düğmeler ve makro bağlantıları, pano grafikleri, korumalar ve
   `.xlsm` kaydı. openpyxl bu üçünü yapamaz.

Böylece kaynak kod düz metin dosyalarında durur — sürüm takibi yapılabilir,
karşılaştırılabilir, gözden geçirilebilir. Değişiklik yapmak için ilgili kaynak
dosya düzenlenir ve `python kur.py` yeniden çalıştırılır.

**VBA modülleri `Import` ile değil `CodeModule.AddFromString` ile aktarılır.**
`Import` dosyayı sistemin ANSI kod sayfasıyla okur; Türkçe karakterler farklı
kod sayfasına sahip bir makinede bozulurdu. `AddFromString` Unicode üzerinden
geçer ve kod sayfasından tamamen bağımsızdır.

**Tekrarlanan bilgi denetlenir.** Sistemde bilinçli olarak dört yerde kopya
vardır: renkler (`tasarim.py` ↔ `modTasarim.bas`), tablo koordinatları
(`uret_yonetim.py` ↔ VBA sabitleri), doğrulama listeleri (`kur.py` ↔
`modModel.bas`) ve **depo sütun düzeni** (`uret_yonetim.py` ↔ `modDepo.bas`).
Sonuncusu en tehlikelisidir: sütunlar sessizce kayarsa gönderim yanlış sütuna
yazılır ve hiçbir şey hata vermez. `test_uretim.py` her başlığın konumunu
karşılık gelen `O_*` / `E_*` sabitiyle tek tek karşılaştırır.

**Yönetim kitabı parolayla kaydedilir.** `kur.py` parolayı `modAyar.bas`'tan
okur; Python tarafında ikinci bir kopya tutulsaydı ikisi sessizce ayrılabilir
ve üretilen kitap, kodun beklediğinden başka bir parolayla şifrelenebilirdi.
Aynı sebeple sayfa koruma parolası da oradan okunur.

**Yeniden üretim veriyi silmez — ama komutu doğru vermek gerekir.**
`python kur.py yonetim --veri <mevcut kitap>` eski kitaptaki `Oneriler` ve
`Olaylar` satırlarını yenisine taşır. Bu bayrak unutulursa yeni dosya boş
gelir; dolu bir kurulumda **yedek klasörü son sığınaktır.**

`kur.py` Excel'in "VBA proje nesne modeline erişime güven" ayarını geçici olarak
açar ve **işi bitince, hata alsa bile, eski haline döndürür**.

---

## 8. Güvenlik — ne gerçek, ne değil

**Tehdit modeli açıkça şudur: sıradan personel.** Amaç, öneri gönderen bir
çalışanın başkalarının önerilerini ve değerlendirme notlarını *kazara ya da
merakla* görmesini engellemektir. Kararlı bir saldırgan, BT yetkisi olan biri
ya da VBA'yı açmayı bilen biri bu sınırların dışındadır ve tasarım onlara göre
yapılmamıştır. Bu, kullanıcının bilerek verdiği bir karardır.

| Katman | Gerçek sınır mı | Kime karşı | Not |
|---|---|---|---|
| Personel / yönetim ekran şifresi | **Hayır** | — | VBA içinde düz durur |
| Sayfa ve kitap koruması | **Hayır** | — | Kazara bozmayı önler |
| **Yönetim kitabının açılış parolası** | **Evet** | Sıradan personel | Dosya AES ile şifrelidir; parolasız açılmaz. Parola VBA'da düz durur |
| NTFS klasör izinleri | **Kısmen** | Sıradan personel | Yalnızca `yedek\` klasörünü korur |

**Neden gizlilik NTFS'ten parolaya indi.** Veri artık Yönetim kitabının içinde
ve personelin oraya *yazması* gerekiyor. Excel bir dosyayı **okumadan
yazamaz**; dolayısıyla personelin o dosya üzerinde okuma hakkı da olmak
zorundadır. Eski "bırakma kutusu" — yaz evet, oku hayır — bu modelde
kurulamaz. Yerine dosyanın kendisi şifrelendi.

**Personel neyi görebilir:** `yonetim\` klasörünü ve içindeki dosya adlarını.
**Neyi göremez:** dosyaların içini. **Neyi yapabilir:** dosyayı kopyalayabilir
(ama açamaz) ve — teknik olarak — silebilir. Silinmeye karşı koruma yedeklerdir.

**Neden izinler dosyaya değil klasöre verilir.** Excel kaydederken dosyayı
yerinde değiştirmez: geçici bir dosya yazıp aslının yerine koyar. Dosyaya
doğrudan verilen açık haklar bu sırada kaybolabilir; klasörden miras alınanlar
kalır. Ayrıca personelin klasörde dosya oluşturma **ve silme** hakkı olmak
zorundadır, çünkü Excel'in geçici dosyası ve `~$` sahiplik dosyası orada
oluşur.

Önerilen izin düzeni (ayrıntısı `KURULUM.md` madde 3):

| Klasör / dosya | Tüm personel | Değerlendirme ekibi |
|---|---|---|
| `projeoneri\` | Okuma + Çalıştırma | Tam denetim |
| `ProjeOneri.xlsm` | **Salt okunur** (yoksa ilk açan kilitler) | Tam denetim |
| `yonetim\` | **Değiştir** (mirasla) | Tam denetim |
| `yonetim\yedek\` | **Hiçbir hak** (miras kesilir) | Tam denetim |

Bu model `testler	est_izinler.py` ile gerçek NTFS izinleri altında sınanır:
klasörler kilitlenir, gerçek gönderim makrosu çalıştırılır ve gönderimin
çalıştığı, kaydetmeden sonra dosyanın hâlâ erişilebilir olduğu (yani izinlerin
kaydetmeyi atlattığı), buna karşılık yedek klasörünün listelenemediği ve
üzerine yazılamadığı tek tek doğrulanır.

Salt okunur açılan bir kitabın form olarak kullanılabilmesi Excel'in
davranışına dayanır: dosya bellekte düzenlenebilir, yalnızca kitabın kendisi
kaydedilemez. Makronun yönetim kitabına yazmasını engellemez.

---

## 9. Test kapsamı

`python testler	um_testler.py` — **250'den fazla kontrol, dört aşamada**,
yaklaşık dört dakika. Bu bir taklit (mock) testi değildir: üretilen `.xlsm`
dosyalarını gerçekten Excel'de açar, VBA makrolarını çalıştırır ve sonuçları
**diskteki gerçek depodan** okur.

| Aşama | Neyi denetler |
|---|---|
| `test_uretim.py` | Excel açmadan: dosyalar üretildi mi, yönetim kitabı gerçekten **şifreli** mi ve bilinen parolayla açılıyor mu, sayfa düzeni ve adlandırılmış aralıklar yerinde mi, **tekrarlanan bilgi ayrışmış mı** (özellikle depo sütun düzeni) |
| `test_uctan_uca.py` | Gerçek Excel'de: gönderim, konsolidasyon, durum türetme, ekle-only geçmiş, göstergeler, form ve değerlendirme ekranlarının **gerçek düğme yolları**, kitabın **gerçek açılışı** (yedek + salt okunura geçiş), kaydetme yasağı, **depo başkasındayken gönderimin anlaşılır hata vermesi** |
| `test_eszamanlilik.py` | Arka arkaya 20 gönderim; 3 ayrı süreç × 3 gönderim **artı** aynı anda değerlendirme kaydeden bir "ekip" süreci |
| `test_izinler.py` | `yonetim\` ve `yedek\` klasörlerini **gerçek NTFS izinleriyle** kilitler; gönderimin çalıştığını, kaydetmenin izinleri bozmadığını, yedek klasörünün kapalı kaldığını doğrular |

Özellikle korunan davranışlar: ikinci değerlendirme ilkinin üzerine yazmaz;
yalnızca durum değiştiren kısmi bir olay önceki karar notunu silmez; sahipsiz
bir olay konsolidasyonu durdurmaz; sıralı numaralar boşluksuz gider ve iki
öneriye aynı numara verilmez; grafikler hiçbir kategoriyi düşürmez; **arkada
görünmez bir Excel süreci kalmaz.**

**Sızıntı kontrolü neden var:** her yazma gizli bir Excel örneği açar.
Kapatılmayan bir örnek dosya kilidini süresiz tutar ve bir sonraki gönderim
otuz saniye bekleyip "başkası kullanıyor" der — oysa kimse kullanmıyordur.
Testler bu yüzden başlangıç ve bitiş `EXCEL.EXE` sayılarını karşılaştırır.

**Testin göremediği — bilinçli sınır:** Testler makroları COM üzerinden
çağırır. Bu yol Excel'in makro güvenlik ayarını, "İçeriği Etkinleştir"
uyarısını, parola sorulmasını ve düğmelere basmayı hiç görmez. Yani testler
"kod doğru mu" sorusunu yanıtlar, "kullanıcı bu dosyayı açtığında ne olur"
sorusunu yanıtlamaz. Kurulumdan sonra `KURULUM.md` madde 7'deki **elle kontrol
listesi şarttır.**

**Görünmez diyalog tuzağı.** Görünmez bir Excel'de açılan herhangi bir iletişim
kutusu — bir `MsgBox`, bir **parola sorusu** ya da VBA'nın kendi çalışma zamanı
hata penceresi — ekranda görünmez ama makro geri dönmez: otomasyon sessizce
sonsuza kadar bekler. Üç önlem alındı: `modUI` bir **sessiz mod** taşır,
`modDepo` kullanıcıyla hiç konuşmaz ve depoyu **her zaman parolasıyla** açar,
test altyapısı da bir **bekçi** çalıştırır (belirli süre dönmeyen makronun
adını yazıp süreci sonlandırır).

**Bekçi bu sürümde gerçek bir kilitlenme yakaladı** ve nedeni öğrenilmeye
değer: `On Error Resume Next` altında bir `Dir$` döngüsü **sonsuza gidebilir.**
`Dir$` hata verdiğinde değer döndürmez; atama yapılmadığı için değişken eski
değerinde kalır ve `Do While Len(ad) > 0` hep doğru olur. Yedek klasörü
personele kapatıldığı anda tam olarak bu oluyordu. Artık her `Dir$` çağrısından
sonra `Err` denetleniyor ve ayrıca bir üst sınır var.

**Tasarım gözden geçirmesi.** `python testler\goruntu_al.py` her ekranı örnek
veriyle doldurup PDF olarak dışa aktarır. Otomatik testler bir ekranın doğru
*çalıştığını* gösterir, iyi *göründüğünü* gösteremez; hizalama, renk ve
yerleşim böyle denetlenir.

---

## 10. Bilinen kısıtlar, riskler ve açık maddeler

**1. Makro izni — sistemi tümden engelleyebilir**
Banka BT'si makroları grup ilkesiyle kapattıysa bu sürüm hiç açılmaz. Kuruma
sorulması gerekenler: (a) makrolar kullanıcı onayıyla çalışabiliyor mu,
(b) ağ paylaşımı "Güvenilir Konum" olarak tanımlanabilir mi? Tanımlanmazsa
sistem yine çalışır, kullanıcı her açılışta *İçeriği Etkinleştir* der.

**2. Tek dosya = tek kırılma noktası**
Bütün veri tek bir dosyadadır. O dosya silinirse, bozulursa ya da yanlışlıkla
üzerine yazılırsa her şey gider. Üç önlem var: günlük yedek, `yedek\`
klasörünün personele kapalı olması ve kitabın kendisini asla kaydetmemesi.
**Yine de kurumun düzenli dosya yedeği alması önerilir.**

**3. Ekip kitabı yazma kipinde kalırsa personel gönderemez**
`ChangeFileAccess` başarısız olursa ya da kullanıcı Excel'in "Yine de Düzenle"
şeridine basarsa kilit ekipte kalır ve gönderimler otuz saniye bekleyip hata
verir. Giriş ekranındaki uyarı bandı bu durumun tek görünür işaretidir; bandı
gören kişi kitabı kapatıp yeniden açmalıdır.

**4. Gizlilik VBA'yı açan birine karşı korumaz**
Depo parolası `modAyar.bas` içinde düz durur ve üretilen kitapların VBA'sında
da öyle. Tehdit modeli sıradan personeldir (madde 8).

**5. Her gönderim tam bir dosya yazımıdır**
Bir satır eklemek dosyanın tamamının yeniden yazılması demektir. Ölçülen süre
yaklaşık 3 saniyedir ve yazmalar sıraya girer: aynı anda gönderim yapan
kişi sayısı kadar bu süre beklenebilir. Sıra bekleme bütçesi 60 saniyedir. Ağ
paylaşımında dosya büyüdükçe bu süre uzar.

**6. Ölçek**
Birkaç bin satıra kadar sorun beklenmez. On binlere çıkılırsa dosya büyür ve
her gönderim yavaşlar; o noktada eski yılların satırları bir arşiv kitabına
taşınmalıdır.

**7. Zorunlu duyarlılık etiketi (MIP) riski**
Kurum Office'te "duyarlılık etiketi seçmeden kaydedilemez" kuralı
uyguluyorsa, kaydetme sırasında çıkan etiket penceresi `DisplayAlerts` ile
bastırılamaz ve **görünmez Excel örneğini kilitler.** BT'ye sorulacak maddeler
arasına eklenmelidir.

**8. Klasör OneDrive/SharePoint ile eşlenmişse**
Eşlenmiş klasörlerde Excel dosyanın konumunu disk yolu olarak değil
`https://...` adresi olarak bildirir; VBA böyle bir adrese yazamaz.
`modDosyaIO.bas` içindeki `YerelYol()` bu adresi diskteki gerçek klasöre
çevirir. Ancak Excel'in **Güvenilir Konumlar listesi disk yollarına göre
çalıştığı için** eşlenmiş klasörlerde güven ayarı beklendiği gibi
davranmayabilir. **Öneri: çıktıları OneDrive altındaki bir klasörden
çalıştırmayın.** Gerçek kurulum bir UNC ağ paylaşımında olacağı için
(`\\sunucu\paylasim\projeoneri`) durum üretimde oluşmaz.

**9. Kişi bazlı yetki yok**
Yönetim erişimi tek ortak şifredir. Değerlendirme kaydında hangi Windows
kullanıcısının yaptığı yazılır, ama şifre paylaşıldığı için bu kimlik doğrulama
değil, yalnızca iz kaydıdır.

**10. Anlık bildirim yok**
Yeni öneri geldiğinde kimseye haber gitmez. Değerlendirme ekibi kitabı açıp
*Önerileri Yenile* demelidir.

**11. Excel sürümü**
Doğrulanan tek hedef Windows masaüstü Excel'dir (Microsoft 365 / 2019+).
Excel for Mac ve Excel for the web makroları çalıştırmaz; bu yaklaşım oralarda
kullanılamaz.

---

## 11. Sürüm 1 ile karşılaştırma

| | Sürüm 1 (Flask sunucu) | Sürüm 2 (iki Excel dosyası) |
|---|---|---|
| Sürekli açık bilgisayar | gerekli | gerekmez |
| Kullanıcıda kurulum | yok (tarayıcı) | yok (Excel) |
| Makro izni | gerekmez | **gerekir** |
| Veri nerede | sunucudaki veritabanı | **Yönetim kitabının içinde** |
| Takip numarası | sıralı `PRJ-2026-000001` | sıralı `PRJ-2026-0001` |
| Eşzamanlı yazma | sunucu sıraya sokar | kısa dosya kilidi + yeniden deneme |
| Anlık bildirim | mümkün | yok |
| Bakım | sunucu izlenmeli | **iki dosya yedeklenmeli** |

Sürüm 1 daha akıcı bir deneyim sunar; Sürüm 2 daha az kurumsal onay gerektirir.
Karar, bankanın sunucu tahsis edip edemeyeceğine ve makro politikasına bağlıdır.

---

## 12. Sürdürme

| İş | Nasıl |
|---|---|
| Şifre / parola / başlık / durum listesi değişikliği | `kaynak\vba\*.bas` ya da `kaynak\tasarim.py` düzenle, `python kur.py` |
| Ekran veya alan değişikliği | `kaynak\uret_*.py` düzenle, `python kur.py` |
| Renk veya tipografi değişikliği | `kaynak\tasarim.py` **ve** `modTasarim.bas` birlikte düzenle (kur.py ikisinin aynı kaldığını denetler) |
| **Dolu bir kurulumu güncelleme** | **`python kur.py yonetim --veri <mevcut ProjeYonetim.xlsm>`** — bayrak unutulursa yeni dosya boş gelir |
| Doğrulama | `python testler\tum_testler.py` + `KURULUM.md` madde 7'deki elle liste |
| Tasarım gözden geçirme | `python testler\goruntu_al.py` → ekranların PDF'i |
| **Yedekleme** | **`yonetim\` klasörünün tamamı** — `ProjeYonetim.xlsm` verinin kendisidir |

**Yedekleme kuralı değişti ve önemlidir.** Eski sürümde çalışma kitapları
üretilebilir dosyalardı, veri değildi; verinin tamamı klasördeki metin
dosyalarındaydı. **Artık öyle değil:** `ProjeYonetim.xlsm` verinin kendisidir
ve kaybı geri alınamaz. `ProjeOneri.xlsm` hâlâ üretilebilir bir dosyadır.

Sistemin kendi günlük yedeği (`yonetim\yedek\`) bir kolaylıktır, kurumsal
yedeğin yerini tutmaz: aynı diskte durur ve yalnızca son yedi günü kapsar.
