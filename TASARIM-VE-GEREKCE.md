# Proje Öneri Sistemi — Sürüm 2 (Ortak Klasör)
## Tasarım, gerekçe ve kısıtlar

Bu belge Sürüm 2'nin **neden böyle tasarlandığını** anlatır. Nasıl kurulacağı
`KURULUM.md`'de, sistemin genel tanıtımı `OKU-BENI.md`'dedir.

---

## 1. Bir bakışta

|  |  |
|---|---|
| **Amaç** | Personelin iyileştirme önerisi göndermesi, Değerlendirme ekibinin bunları PDCA döngüsüyle takip etmesi |
| **Yaklaşım** | Sunucu yok. Her şey bir ağ paylaşımındaki Excel dosyaları ve klasörlerle yürür |
| **Kullanıcıda kurulum** | Yok. Herkeste zaten olan Excel yeterli |
| **Tek gerçek gereksinim** | Makroların çalışmasına izin verilmesi |
| **Durum** | Tamamlandı; 145'ten fazla otomatik kontrolle, gerçek Excel'de doğrulandı |
| **Karar bekleyen** | Banka BT'sinin makro ve güvenilir konum politikası (madde 9) |

**Neden bu sürüm var:** Sürüm 1 sürekli açık bir bilgisayarda Flask sunucusu
çalıştırır. Böyle bir makine tahsis edilemezse ya da BT ağda uygulama sunucusu
istemezse tüm sistem çöker. Sürüm 2 aynı işi **hiçbir sunucu olmadan** yapar;
iki sürüm birbirinden bağımsızdır, hangisi uygunsa o seçilir.

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
> Ortak klasör yalnızca bir dosya deposudur; üzerinde çalışan bir program yoktur.

### Kısıt 2 — `file://` ile açılan HTML sayfası diske yazamaz

İlk akla gelen çözüm Sürüm 1'in HTML sayfalarını olduğu gibi ortak klasöre
koymaktı. Çalışmaz: tarayıcılar `file://` protokolüyle açılan bir sayfanın
diske yazmasını güvenlik gereği engeller. Sunucu olmadan HTML'in yazma yolu yok.

> **Sonuç:** Excel + VBA seçildi. Excel yazma yetkisine sahiptir, herkeste
> kuruludur ve banka personeli zaten Excel biliyor.

### Kısıt 3 — Aynı dosyaya iki kişi aynı anda yazamaz

Ortak bir `oneriler.xlsx` dosyası olsaydı, ikinci kullanıcı "dosya kilitli"
uyarısı alırdı. Kilit bekleme, yeniden deneme, kuyruk gibi çözümler sunucusuz
ortamda güvenilir değildir.

> **Sonuç — sistemin temel kuralı:** *Kimse ortak bir dosyaya yazmaz.*
> Her gönderim `yonetim\oneriler\` klasörüne **kendi dosyasını** bırakır. Bir kullanıcı
> yalnızca kendi oluşturduğu dosyaya dokunur. Çakışma önlenmiş değil,
> **yapısal olarak imkânsız** hale getirilmiştir.

### Kısıt 4 — VBA içindeki şifre gerçek bir güvenlik sınırı değildir

Şifreler kodun içinde durur. VBA proje parolaları ücretsiz araçlarla kırılabilir.

> **Sonuç:** Şifreler yalnızca "yanlış ekrana yanlışlıkla girmeyi" engeller.
> Asıl erişim denetimi **ağ klasörünün NTFS izinleridir**.

---

## 4. Nasıl çalışır

```
                 PERSONEL                              KAİZEN EKİBİ
            ProjeOneri.xlsm                      yonetim\ProjeYonetim.xlsm
                    │                                        │
         şifre → form → Gönder                      şifre → Önerileri Yenile
                    │                                        │
                    ▼                                        │
        ┌─────────────────────────┐                          │
        │ yonetim\oneriler\<yıl>\ │ ───────── okur ─────────►│
        │  PRJ-26A7K.txt          │                          │
        │  BIRAKMA KUTUSU:        │                          ▼
        │  yaz evet, oku hayır    │                 değerlendirme yapılır
        └─────────────────────────┘
                                                             │
                                                             ▼
                                            ┌────────────────────────┐
                             okur ◄──────── │  degerlendirme\        │
                                            │   her değişiklik = yeni│
                                            │   dosya (ekle-only)    │
                                            └────────────────────────┘
```

**`yonetim\oneriler\` bir bırakma kutusudur.** Personel oraya yazabilir ama
içini göremez: listeleyemez, kimsenin önerisini okuyamaz, hiçbir dosyayı
silemez (NTFS izinleri; `KURULUM.md` madde 3). Klasör `yonetim\` içindedir;
personele o klasör üzerinde yalnızca *geçiş* hakkı verilir, listeleme hakkı
verilmez. Böylece `projeoneri\` altında görünen tek şey çalışma kitabı ve
`yonetim\` adıdır; değerlendirme notlarına erişim yoktur.

**Ayrı bir "gelen kutusu" yoktur.** Ara bir tasarımda öneriler önce
`gelen\` klasörüne bırakılıp okunduktan sonra arşive taşınıyordu. Taşımanın
tek faydası, izinler yanlış kurulursa açığa çıkacak dosya sayısını
sınırlamaktı; buna karşılık ikinci bir klasör, bir taşıma adımı ve yeni bir
hata yolu getiriyordu. İzinler doğru kurulduğunda iki tasarım arasında fark
olmadığı için tek klasörde birleştirildi: dosya baştan itibaren kalıcı
yerindedir, sonradan hiçbir yere taşınmaz.

**Yazma tek adımlıdır ve iki güvence verir.** Dosya doğrudan son adıyla,
“varsa oluşturma” kipinde açılır:

1. *Üzerine yazmaz.* Aynı adda bir dosya varsa işlem başarısız olur ve mevcut
   dosyaya dokunulmaz; gönderim yeni bir numarayla yeniden denenir.
2. *Yarım dosya okunmaz.* Kaydın sonuna bir `kayit_sonu=1` satırı yazılır.
   Yazma yarıda kesilirse bu satır oluşmaz ve okuyucu dosyayı yok sayar.

> **Neden `.tmp` + yeniden adlandırma değil?** İlk tasarım önce `.tmp` yazıp
> sonra yeniden adlandırıyordu. Yeniden adlandırma, kaynak dosya üzerinde
> **silme** yetkisi ister — oysa bırakma kutusunda personelin silme yetkisi
> yoktur ve işletim sistemi işlemi reddeder. Bu, denenerek görüldü
> (`testler\test_izinler.py`). Dosyayı doğrudan son adıyla oluşturmak aynı iki
> güvenceyi yalnızca **yazma** yetkisiyle sağlar.

**Kayıt biçimi** — düz metin, `anahtar=değer`, UTF-8 (BOM'suz):

```
sema=3
oneri_no=PRJ-26A7K
tarih=2026-09-04T10:11:51
ad_soyad=...
sicil_no=...
mevcut_durum=...
oneri_basligi=...
cozum_onerisi=...
beklenen_fayda=...
gonderen_bilgisayar=...
gonderen_kullanici=...
```

Excel değil düz metin seçildi: bozulmaya karşı dayanıklı, Not Defteri'yle
okunabilir, kilitlenmez ve gerekirse elle onarılabilir.

**Gönderim kaydında `durum` alanı yoktur.** Gönderim dosyası değişmezdir;
durum yalnızca değerlendirme olaylarından türetilir. Böylece "tek doğruluk
kaynağı" sorusu ortadan kalkar: aynı bilgi iki yerde tutulmaz.

**Çok satırlı alanlar** tek satıra sığdırılır: değer içindeki satır sonları
sabit bir belirtece (`<|>`) çevrilir, okunurken geri açılır. Böylece her alan
dosyada tam olarak bir satır kaplar ve ayrıştırma basit kalır.

**Kodlama.** Tüm okuma/yazma `ADODB.Stream` üzerinden UTF-8 yapılır. VBA'nın
yerleşik `Open/Print` komutu dosyayı ANSI yazar; Türkçe karakterler farklı kod
sayfasına sahip bir makinede bozulurdu.

**Öneri numarası** `PRJ-YYXXX` biçimindedir; örnek: `PRJ-26A7K`.
Dokuz karakterdir — telefonda söylenebilsin, elle yazılabilsin, bir kenara not
edilebilsin diye. Rastgele ekin alfabesinden `0/O` ve `1/I` çıkarılmıştır.

Sıralı numara (`000001`) kullanılamaz — sıralı sayaç ortak bir dosya
gerektirir, o da Kısıt 3'e takılır. Tarih + rastgele ek, merkezi bir sayaç
olmadan benzersizliği sağlar. Numara kısaldığı için aynı numaranın iki kez
üretilme ihtimali doğar; bu yüzden **yazma işlemi hiçbir dosyanın üzerine
yazmaz.** Hedef dosya varsa yazma başarısız olur ve gönderim yeni bir
numarayla yeniden denenir. Böylece kısa numara, kayıt kaybı riski getirmeden
kullanılabilir. Dosya adı öneri numarasının kendisidir.

**Olay dosyalarının adı 10 ms çözünürlüklü zaman damgası taşır**
(`<oneri_no>_<YYYYMMDDHHMMSSss>_<rastgele>.txt`). Olayların sırası
adlarından okunduğu için saniye çözünürlüğü yetmezdi: aynı saniyede yazılan
iki olayın sırasını rastgele ek belirler ve durum yanlış türetilebilirdi.
(Gönderim numarası artık zaman sıralı değildir; liste sıralaması dosya adına
değil kayıttaki `tarih` alanına bakar.)

**Yıl alt klasörleri** (`oneriler\2026\`) ilk günden kullanılır; klasör başına
dosya sayısı düşük kalır ve yıllık arşivleme doğal olur.

**Yönetim kitabı tek doğruluk kaynağı değildir.** Ekrandaki her şey `oneriler\` ve
`degerlendirme\` klasörlerinden yeniden üretilebilir. Kitap silinse, bozulsa
veya yeniden kurulsa veri kaybolmaz.

---

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
| `modAyar.bas` | Klasör yolları, şifre sabitleri, şema sürümü |
| `modTasarim.bas` | Renk ve tipografi sabitleri, durum rozetleri |
| `modDosyaIO.bas` | UTF-8 okuma/yazma, atomik yazma, klasör tarama, OneDrive yol çevirisi |
| `modModel.bas` | Değerlendirme modeli: durumlar, PDCA, durum anlamları |
| `modUI.bas` | Ekran yönetimi, şifre kapısı, koruma, mesajlar, sessiz mod |
| `modGonderim.bas` | Personel tarafı: form doğrulama, öneri no, gönderim |
| `modKonsolide.bas` | Klasörleri okuyup olayları tekrar oynatarak tabloyu kurma |
| `modDegerlendirme.bas` | Değerlendirme ekranı, ekle-only olay kaydı, geçmiş |
| `modPano.bas` | Göstergeler ve grafik kaynak verisi |

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

**Tekrarlanan bilgi denetlenir.** Sistemde bilinçli olarak üç yerde kopya
vardır: renkler (`tasarim.py` ↔ `modTasarim.bas`), tablo koordinatları
(`uret_yonetim.py` ↔ VBA sabitleri) ve doğrulama listeleri (`kur.py` ↔
`modModel.bas`). Bu kopyaların sessizce birbirinden ayrılması fark edilmesi zor
hatalara yol açardı; `kur.py` renkleri her üretimde, `test_uretim.py` üçünü
birden karşılaştırır.

`kur.py` Excel'in "VBA proje nesne modeline erişime güven" ayarını geçici olarak
açar ve **işi bitince, hata alsa bile, eski haline döndürür**.

---

## 8. Güvenlik — ne gerçek, ne değil

| Katman | Gerçek sınır mı | Not |
|---|---|---|
| Personel şifresi | **Hayır** | VBA içinde düz durur |
| Yönetim şifresi | **Hayır** | VBA içinde düz durur |
| Sayfa/kitap koruması | **Hayır** | Kazara bozmayı önler |
| NTFS klasör izinleri | **Evet** | Asıl erişim denetimi budur |

Artık `yonetim\` altından dışarı bir dosya üretilmez: rapor ekranı kaldırıldığı
için sistemin ürettiği her şey izinlerle korunan klasörlerin içinde kalır
(gerekçesi 5. bölümde).

Önerilen izin düzeni (ayrıntısı `KURULUM.md` madde 3):

| Klasör | Kimler | İzin |
|---|---|---|
| `projeoneri\` | Tüm personel | Okuma + Çalıştırma |
| `ProjeOneri.xlsm` | Tüm personel | **Salt okunur** (yoksa ilk açan kilitler) |
| `yonetim\` | Tüm personel | **Yalnızca geçiş** — listeleme yok, miras yok |
| `yonetim\oneriler\` | Tüm personel | **Yalnızca yazma** — listeleme, okuma, silme yok |
| `yonetim\` içindeki diğer her şey | Yalnızca Değerlendirme ekibi | Tam denetim |

Bırakma kutusunun **bırakma kutusu** olması kritiktir. Yalnızca silmeyi
engellemek yetmez: listeleme ve okuma da kapatılmalıdır, aksi halde herkes
herkesin önerisini okuyabilir.

`yonetim\` üzerindeki geçiş hakkının **miras bayrağı yoktur**; yalnızca o
klasörün kendisine uygulanır, kardeş klasörlere sızmaz. Yanlış kurulsa bile
zarar sınırlıdır: geçiş hakkı bir dosyanın içeriğini okumaya yetmez, bunun
için ayrıca *veri okuma* hakkı gerekir.

**Komut sırası kurulumun en kırılgan noktasıdır:** önce alt klasörler
ayarlanmalı, `yonetim\` en son kısıtlanmalıdır. Ters sırada izin komutlarının
kendisi çalışamaz hale gelir ve klasörler sessizce erişilemez kalır — bu
tasarlanırken bizzat karşılaşılan bir hatadır, `KURULUM.md`'de komutlar
sırasıyla verilmiştir.

Bu izin modeli `testler\test_izinler.py` ile gerçek NTFS izinleri altında
sınanır: klasörler kilitlenir, gerçek gönderim makrosu çalıştırılır ve
gönderimin çalıştığı, buna karşılık yönetim klasörünün listelenemediği,
değerlendirme notlarının okunamadığı ve üzerlerine yazılamadığı, yönetim
kitabının okunamadığı, bırakılan önerinin okunamadığı ve silinemediği tek tek
doğrulanır.

Salt okunur açılan bir kitabın form olarak kullanılabilmesi Excel'in
davranışına dayanır: dosya bellekte düzenlenebilir, yalnızca kitabın kendisi
kaydedilemez. Makronun ortak klasöre yazmasını engellemez.

---

## 9. Test kapsamı

`python testler\tum_testler.py` — **145'ten fazla kontrol, üç aşamada.**
Bu bir taklit (mock) testi değildir: üretilen `.xlsm` dosyalarını gerçekten
Excel'de açar, VBA makrolarını çalıştırır ve sonuçları dosya sisteminden
doğrular.

| Aşama | Neyi denetler |
|---|---|
| `test_uretim.py` | Excel açmadan: dosyalar üretildi mi, VBA projesi var mı, sayfa düzeni ve adlandırılmış aralıklar yerinde mi, **tekrarlanan bilgi ayrışmış mı** |
| `test_uctan_uca.py` | Gerçek Excel'de: gönderim, konsolidasyon, durum türetme, ekle-only geçmiş, göstergeler, form ve değerlendirme ekranlarının **gerçek düğme yolları** |
| `test_eszamanlilik.py` | Aynı saniyede 30 gönderim; 3 ayrı süreç ve 3 ayrı Excel örneğiyle paralel gönderim |
| `test_izinler.py` | `yonetim\` ve `yonetim\oneriler\` klasörlerini **gerçek NTFS izinleriyle** kilitler; gönderimin çalıştığını, buna karşılık listeleme/okuma/silmenin engellendiğini doğrular |

Özellikle korunan davranışlar: ikinci değerlendirme ilkinin üzerine yazmaz;
yalnızca durum değiştiren kısmi bir olay önceki karar notunu silmez; yarım
yazılmış dosya okunmaz; grafikler hiçbir kategoriyi düşürmez.

**Testin göremediği — bilinçli sınır:** Testler makroları COM üzerinden çağırır.
Bu yol Excel'in makro güvenlik ayarını, "İçeriği Etkinleştir" uyarısını ve
düğmelere basmayı hiç görmez. Yani testler "kod doğru mu" sorusunu yanıtlar,
"kullanıcı bu dosyayı açtığında ne olur" sorusunu yanıtlamaz. Kurulumdan sonra
`KURULUM.md` madde 7'deki **elle kontrol listesi şarttır.**

**Görünmez diyalog tuzağı.** Görünmez bir Excel'de açılan herhangi bir iletişim
kutusu — bir `MsgBox`, bir parola sorusu ya da VBA'nın kendi çalışma zamanı
hata penceresi — ekranda görünmez ama makro geri dönmez: otomasyon sessizce
sonsuza kadar bekler. İki önlem alındı: `modUI` bir **sessiz mod** taşır
(mesajlar gösterilmek yerine kaydedilir, testler `SonMesaj()` ile okur) ve
test altyapısı bir **bekçi** çalıştırır (belirli süre dönmeyen makronun adını
yazıp süreci sonlandırır). Geliştirme sırasında karşılaşılan üç ayrı kilitlenme
bu yolla adı konmuş hatalara dönüştü.

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
Tanımlanırsa hiç uyarı görmez.

**2. Klasör OneDrive/SharePoint ile eşlenmişse**
Eşlenmiş klasörlerde Excel dosyanın konumunu disk yolu olarak değil
`https://...` adresi olarak bildirir; VBA böyle bir adrese yazamaz.
`modDosyaIO.bas` içindeki `YerelYol()` bu adresi diskteki gerçek klasöre
çevirir. Ancak Excel'in **Güvenilir Konumlar listesi disk yollarına göre
çalıştığı için** eşlenmiş klasörlerde güven ayarı beklendiği gibi
davranmayabilir. **Öneri: çıktıları OneDrive altındaki bir klasörden
çalıştırmayın.** Testler bu belirsizliği dışarıda bırakmak için `%TEMP%`
altında — OneDrive dışında — kendi geçici paylaşımını kurar. Gerçek kurulum bir
UNC ağ paylaşımında olacağı için (`\\sunucu\paylasim\projeoneri`) durum üretimde
oluşmaz.

**3. Kişi bazlı yetki yok**
Yönetim erişimi tek ortak şifredir. Değerlendirme kaydında hangi Windows
kullanıcısının yaptığı yazılır, ama şifre paylaşıldığı için bu kimlik doğrulama
değil, yalnızca iz kaydıdır.

**4. Ölçek**
`oneriler\` klasörü her yenilemede baştan taranır. Yıl alt klasörleri sayesinde
birkaç bin dosyaya kadar sorunsuzdur; on binlere çıkılırsa eski yıl
klasörlerini arşive taşımak yeterlidir.

**5. Anlık bildirim yok**
Yeni öneri geldiğinde kimseye haber gitmez. Değerlendirme ekibi kitabı açıp
*Önerileri Yenile* demelidir.

**6. Excel sürümü**
Doğrulanan tek hedef Windows masaüstü Excel'dir (Microsoft 365 / 2019+).
Excel for Mac ve Excel for the web makroları çalıştırmaz; bu yaklaşım oralarda
kullanılamaz.

**7. Bu sürümde kapatılan açık madde**
Önceki sürümde "yönetim kitabındaki düğmeler makro hatası veriyor" diye açık
bir madde vardı ve nedeninin OneDrive olduğu değerlendiriliyordu. Bu sürümde
neden bulundu ve giderildi: `ThisWorkbook` modülündeki bir olay imzası yanlış
yazılmıştı (`Cancel` parametresi `ByVal` olarak tanımlanmıştı; doğrusu `ByRef`).
Böyle bir hata **derleme** hatasıdır ve `ThisWorkbook` modülü ancak ilk olay
tetiklendiğinde derlendiği için hiçbir yerde ortaya çıkmaz — ne üretimde, ne
makroları tek tek çalıştıran testlerde. Kullanıcı ilk kez bir hücreye
yazdığında ya da çift tıkladığında ham bir Visual Basic hata penceresi olarak
çıkar. Artık iki kitabın `ThisWorkbook` modülü de bir `DerlemeSinamasi()`
işlevi taşır ve testler bunu çağırarak modülü derlenmeye zorlar.

---

## 11. Sürüm 1 ile karşılaştırma

| | Sürüm 1 (Flask sunucu) | Sürüm 2 (ortak klasör) |
|---|---|---|
| Sürekli açık bilgisayar | gerekli | gerekmez |
| Kullanıcıda kurulum | yok (tarayıcı) | yok (Excel) |
| Makro izni | gerekmez | **gerekir** |
| Takip numarası | sıralı `PRJ-2026-000001` | yıl + rastgele `PRJ-26A7K` |
| Eşzamanlı yazma | sunucu sıraya sokar | yapısal olarak imkânsız |
| Anlık bildirim | mümkün | yok |
| Bakım | sunucu izlenmeli | dosya klasörü yedeklenmeli |

Sürüm 1 daha akıcı bir deneyim sunar; Sürüm 2 daha az kurumsal onay gerektirir.
Karar, bankanın sunucu tahsis edip edemeyeceğine ve makro politikasına bağlıdır.

---

## 12. Sürdürme

| İş | Nasıl |
|---|---|
| Şifre / başlık / durum listesi değişikliği | `kaynak\vba\*.bas` ya da `kaynak\tasarim.py` düzenle, `python kur.py` |
| Ekran veya alan değişikliği | `kaynak\uret_*.py` düzenle, `python kur.py` |
| Renk veya tipografi değişikliği | `kaynak\tasarim.py` **ve** `modTasarim.bas` birlikte düzenle (kur.py ikisinin aynı kaldığını denetler) |
| Doğrulama | `python testler\tum_testler.py` + `KURULUM.md` madde 7'deki elle liste |
| Tasarım gözden geçirme | `python testler\goruntu_al.py` → ekranların PDF'i |
| Yedekleme | **Yedeklenmesi gereken `yonetim\` klasörünün tamamıdır** (öneriler ve değerlendirmeler orada). `.xlsm` dosyaları her zaman `kur.py` ile yeniden üretilebilir |

Yedekleme kuralı önemlidir: çalışma kitapları üretilebilir dosyalardır, veri
değildir. Verinin tamamı iki klasördeki düz metin dosyalarındadır.
