# XJ Birimi — Proje Öneri Formu

Personel iyileştirme önerisi gönderir; Değerlendirme ekibi bunları PDCA döngüsüyle
değerlendirir, önceliklendirir ve panodan izler.

**Sunucu yoktur, veritabanı yoktur, kayıt dosyası yoktur.** Sistem bir ağ
paylaşımındaki **iki Excel dosyasından** ibarettir. Kullanıcıda kurulum
gerekmez; herkeste zaten olan Excel yeterlidir. Tek gerçek gereksinim
makroların çalışmasına izin verilmesidir.

| | |
|---|---|
| Ekranlar | Öneri formu · Pano · Liste · Değerlendirme |
| Veri | Yönetim kitabının içinde, iki gizli sayfada |
| Öneri no | `PRJ-2026-0001` — sıralı, kısa, telefonda söylenebilir |
| Gizlilik | Yönetim kitabı açılış parolasıyla şifreli |
| Kurulum | Kullanıcıda yok; paylaşıma bir kez kopyalama |
| Doğrulama | 250+ otomatik kontrol, gerçek Excel'de |

---

## Nasıl çalışır

```
        PERSONEL                                  KAİZEN EKİBİ
   ProjeOneri.xlsm                         yonetim\ProjeYonetim.xlsm
   (salt okunur, şifresiz)                 (parolalı — VERİ BURADA)
          │                                            │
   şifre → form → Gönder                    parola → Pano → Liste
          │                                            │
          │  kilidi alır, gizli bir Excel örneği       │  açılışta:
          │  dosyayı açar, bir satır ekler,            │   yedek al,
          ▼  kaydeder, kapatır, bırakır  (~3 sn)       │   SALT OKUNURA GEÇ
  ┌──────────────────────────────┐ ◄───────────────────┘
  │  ProjeYonetim.xlsm           │   Yenile → diskten çeker
  │    Oneriler  (değişmez)      │   Kaydet → yeni olay satırı
  │    Olaylar   (ekle-only)     │
  └──────────────────────────────┘
          │
          ▼
  yonetim\yedek\ProjeYonetim_YYYYMMDD.xlsm     günde bir, son 7
```

Üç kural sistemin belkemiğidir:

1. **Herkes aynı dosyaya, kısa ve dışlayıcı bir işlemle yazar.** Kilidi al,
   aç, bir satır ekle, kaydet, kapat, kilidi bırak. Sıra bekleyen geri
   çekilip yeniden dener; bütçe dolarsa kullanıcı anlaşılır bir hata görür ve
   **yazdıkları formda durur.**

   Kilit, dosyanın yanında bir an için oluşturulan `ProjeYonetim.xlsm.kilit`
   dosyasıdır. Excel'in kendi kilidi **yetmez**: Excel kaydederken dosyayı
   geçici bir kopyayla değiştirir ve o anda kilidi kısa süreliğine bırakır;
   o aralıkta giren ikinci bir yazıcı, eski içeriği görüp birincinin satırını
   ezebiliyordu. Eşzamanlılık testi bunu ölçtü.
2. **Ekip kitabı kilidi tutmaz.** Yönetim kitabı açılır açılmaz kendini salt
   okunur kipe alır. Bunsuz, ekip kitabı gün boyu açık tuttuğunda hiç kimse
   öneri gönderemezdi. Kitap ayrıca **asla kaydedilmez**: bellekteki eski bir
   kopyanın diskin üzerine yazılması, o gün gelen bütün önerileri silerdi.
3. **Geçmiş silinmez.** Her değerlendirme `Olaylar` sayfasının sonuna yeni bir
   satır ekler. Bir önerinin güncel durumu, gönderim satırının üzerine kendi
   olaylarının sırayla uygulanmasıyla bulunur. Silinebilecek bir "geçmiş
   alanı" yoktur.

Bunun sonucu: **ekranda görünen hiçbir şey elle girilmez.** Liste de pano da
her yenilemede o iki sayfadan baştan türetilir.

**Bedeli açıkça:** veri tek bir dosyadadır. O dosyanın silinmesi ya da
bozulması her şeyin kaybı demektir. Günlük yedek bu yüzden vardır ve kurumsal
yedek yine de şarttır.

---

## Değerlendirme modeli

**Durum akışı (PDCA)**

| Durum | PDCA | Anlamı |
|---|---|---|
| Yeni | — | Henüz incelenmedi |
| Değerlendirmede | Planla | Değerlendirme ekibi inceliyor |
| Planlandı | Planla | Kabul edildi, uygulama planlanıyor |
| Pilot Uygulamada | Uygula | Sınırlı alanda deneniyor |
| Ölçümleniyor | Kontrol | Pilot sonuçları ölçülüyor |
| Standartlaştırıldı | Önlem | Yaygınlaştırıldı, süreç oldu |
| Beklemede / Reddedildi | — | Şartlar oluşmadı / gerekçesiyle kapatıldı |

**Değerlendirme girdisi iki alandır:** yeni durum ve karar notu / yorum. Puan,
etki-efor ya da tasarruf rakamı istenmez — değerlendirme, durumu ilerletmek ve
gerekçeyi yazmaktan ibarettir. Her kayıt geçmişe eklenir; reddedilen bir öneri
için gerekçe zorunludur.

**Pano** dört sayı ve iki grafik gösterir: toplam öneri, değerlendirme bekleyen,
uygulamaya geçmiş, bu ay gelen; durum dağılımı ve son 12 ayın gönderim trendi.

---

## Kaynak düzeni

```
kur.py                    iki kitabı sıfırdan üreten betik
kaynak\
  tasarim.py              tasarım sistemi (renk, tipografi, ölçü)
  uret_ortak.py           masthead, kart, form alanı, tablo desenleri
  uret_oneri.py           ProjeOneri.xlsm sayfa düzeni
  uret_yonetim.py         ProjeYonetim.xlsm sayfa düzeni + depo sayfaları + grafikler
  com_kurulum.py          VBA enjeksiyonu, düğmeler, koruma, parolalı .xlsm kaydı
  vba\                    10 modül + iki ThisWorkbook dosyası
testler\
  tum_testler.py          hepsini sırayla çalıştırır (~4 dakika)
  test_uretim.py          Excel açmadan yapısal ve tutarlılık kontrolleri
  test_uctan_uca.py       gerçek Excel'de gerçek makrolar
  test_eszamanlilik.py    paralel süreçlerle çakışma denemesi
  test_izinler.py         gerçek NTFS izinleriyle doğrulama
  goruntu_al.py           ekranları PDF olarak dışa aktarır (tasarım incelemesi)
cikti\                    üretilen .xlsm dosyaları
```

Çalışma kitapları elle hazırlanmaz. `python kur.py` sayfaları,
biçimlendirmeyi, doğrulama listelerini, düğmeleri ve grafikleri kurar; VBA
modüllerini içeri aktarır ve yönetim kitabını parolayla kaydeder. Kaynak kod
düz metin dosyalarında durduğu için sürüm takibi yapılabilir,
karşılaştırılabilir, gözden geçirilebilir.

> **Dolu bir kurulumu güncellerken** `python kur.py yonetim --veri <mevcut
> ProjeYonetim.xlsm>` kullanın. Yönetim kitabı veri deposu olduğu için, bayrak
> unutulursa yeni dosya **boş** gelir.

Verinin durduğu iki sayfa `modDepo.bas`'taki `O_*` ve `E_*` sabitleriyle
`uret_yonetim.py`'deki başlık listelerinin **aynı** olmasına dayanır. Sütunlar
sessizce kayarsa gönderim yanlış sütuna yazılır ve hiçbir şey hata vermez;
`test_uretim.py` her başlığın konumunu tek tek karşılaştırır.

---

## Teknik notlar

Aşağıdakiler tasarımı belirleyen, kolayca gözden kaçan noktalardır:

- **Çakışmanın işareti hata değildir.** Dosya başkası tarafından yazma kipinde
  açıkken `Workbooks.Open(..., Notify:=False)`, uyarılar kapalıyken hata
  vermez: dosyayı **sessizce salt okunur açar.** Yalnızca hataya bakan bir kod
  yazdığını sanardı; `modDepo` açılan kitabın `ReadOnly` özelliğine bakar.
  `Notify` varsayılan değeriyle (`True`) bırakılırsa Excel dosyayı bildirim
  kuyruğuna ekleyip hata verir; salt okunur açmayı bile denemez.
- **`Save`'in başarısı varsayılamaz.** Bir kaydetme hatası yutulursa satır
  diske inmediği halde çağırana başarı döner ve sonraki yazıcı aynı sıra
  numarasını üretir: bir öneri kaybolur, bir numara iki kez verilir.
- **Kilit, Excel örneğinden önce alınır.** Sırasını bekleyen bir yazıcı boşta
  duran gizli bir Excel tutarsa, dört kişi aynı anda gönderdiğinde makine
  onlarca Excel süreciyle tıkanır.
- **Parola her açışta verilmelidir.** Parolası verilmeden açılan şifreli bir
  kitap, görünmez bir Excel'de parola penceresi açar; pencere ekranda görünmez
  ama çağrı geri dönmez. Ölçüldü: 40 saniye sonra hâlâ bekliyordu.
- **Gizli Excel örneği her hata yolunda kapatılmalıdır.** Kapatılmayan bir
  örnek dosya kilidini süresiz tutar ve bir sonraki gönderim otuz saniye
  bekleyip "başkası kullanıyor" der — oysa kimse kullanmıyordur. Örneğin
  gerçekten ölmesi için o örneğe ait **her** nesne referansının bırakılmış
  olması gerekir; bir açık sayfa değişkeni bile yeter.
- **`On Error Resume Next` altında `Dir$` döngüsü sonsuza gidebilir.** `Dir$`
  hata verdiğinde değer döndürmez; atama yapılmadığı için değişken eski
  değerinde kalır ve `Do While Len(ad) > 0` hep doğru olur. Yedek klasörü
  personele kapatıldığı anda tam olarak bu oluyordu.
- **Depo sütunları metin biçimlidir.** Aksi halde Excel `10045` sicil
  numarasını sayıya, ISO tarihi tarihe çevirir; baştaki sıfırlar ve saniye
  bilgisi sessizce kaybolur.
- **Tek hücrelik aralık dizi döndürmez.** `Range.Value` tek hücrede değerin
  kendisini verir; `veri(i, 1)` çağrısı "Type mismatch" ile patlar. Bu yalnızca
  depoda **tam bir** satır varken olur: ilk gönderim çalışır, ikincisi patlar.
- **VBA kaynak kodu.** Modüller `VBComponents.Import` ile değil,
  `CodeModule.AddFromString` ile aktarılır. Import dosyayı sistemin ANSI kod
  sayfasıyla okur; AddFromString Unicode üzerinden geçer.
- **Birleşik hücreler.** Ekran alanlarının çoğu birleşiktir ve adlandırılmış
  aralık yalnızca sol üst hücreyi gösterir. Excel, birleşik bir hücrenin
  *parçası* üzerinde `ClearContents` gibi işlemleri reddeder; bu yüzden kod
  `MergeArea` üzerinden çalışır.
- **Grafik serileri elle kurulur.** `SetSourceData`'nın otomatik tahmini
  güvenilir değildir: bir aralıkta ilk veri satırını başlık sanıp kategoriyi
  düşürür, diğerinde etiket sütununu ikinci bir seri sanar.
- **Hata işleyicileri hata vermemelidir.** VBA'da bir hata işleyicisinin
  içinde oluşan hata yakalanamaz; makro çöker ve kullanıcıya ham bir Visual
  Basic penceresi açılır.

---

## Bilinen sınırlar

1. **Makro izni sistemi tümden engelleyebilir.** Kurum makroları grup ilkesiyle
   kapattıysa bu yaklaşım hiç çalışmaz (`KURULUM.md` madde 4).
2. **Tek dosya, tek kırılma noktası.** Bütün veri `ProjeYonetim.xlsm`
   içindedir. Günlük yedek bir kolaylıktır; kurumsal yedek yine de şarttır.
3. **Gizlilik VBA'yı açan birine karşı korumaz.** Depo parolası kaynak kodda
   düz durur. Tehdit modeli sıradan personeldir.
4. **Ekip kitabı yazma kipinde kalırsa personel gönderemez.** Giriş ekranındaki
   uyarı bandı bunun tek görünür işaretidir; kitabı kapatıp yeniden açmak
   yeterlidir.
5. **Her gönderim tam bir dosya yazımıdır** (~3 saniye) ve yazmalar sıraya
   girer. Dosya büyüdükçe
   uzar; on binlerce satırda arşivleme gerekir.
6. **Kişi bazlı yetki yok.** Yönetim erişimi tek ortak şifredir.
7. **Anlık bildirim yok.** Değerlendirme ekibi kitabı açıp *Önerileri Yenile*
   demelidir.

---

## Başlarken

```
pip install openpyxl pywin32 msoffcrypto-tool
python kur.py
python testler\tum_testler.py
```

Ardından `KURULUM.md`.
