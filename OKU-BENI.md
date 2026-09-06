# XJ Birimi — Proje Öneri Formu

Personel iyileştirme önerisi gönderir; Kaizen ekibi bunları PDCA döngüsüyle
değerlendirir, önceliklendirir ve yönetime parolalı bir rapor üretir.

**Sunucu yoktur.** Her şey bir ağ paylaşımındaki iki Excel dosyası ve birkaç
klasörle yürür. Kullanıcıda kurulum gerekmez; herkeste zaten olan Excel yeterlidir.
Tek gerçek gereksinim makroların çalışmasına izin verilmesidir.

| | |
|---|---|
| Ekranlar | Öneri formu · Konsol · Değerlendirme · Pano · Rapor |
| Veri | Ortak klasörde düz metin dosyaları (UTF-8) |
| Öneri no | `ON-260904-A7K` — kısa, telefonda söylenebilir |
| Kurulum | Kullanıcıda yok; paylaşıma bir kez kopyalama |
| Doğrulama | 145+ otomatik kontrol, gerçek Excel'de |

---

## Nasıl çalışır

```
        PERSONEL                                    KAİZEN EKİBİ
   KaizenOneri.xlsm                          yonetim\KaizenYonetim.xlsm
          │                                              │
   şifre → form → Gönder                    şifre → Önerileri Yenile
          │                                              │
          ▼                                              │
  ┌────────────────────────┐                             │
  │ yonetim\oneriler\<yıl>\│ ────────── okur ───────────►│
  │  ON-260904-A7K.txt     │                             ▼
  │  BIRAKMA KUTUSU:       │                      değerlendirme
  │  yaz evet, oku hayır   │                             │
  └────────────────────────┘                             ▼
                                        ┌────────────────────────────┐
                          okur ◄─────── │ yonetim\degerlendirme\<yıl>│
                                        │  her değişiklik = yeni     │
                                        │  dosya (ekle-only geçmiş)  │
                                        └────────────────────────────┘
                                                         │
                                                         ▼
                                            yonetim\rapor\ → AES parolalı .xlsx
```

Üç kural sistemin belkemiğidir:

1. **Kimse ortak bir dosyaya yazmaz.** Her gönderim `yonetim\oneriler\` altına
   kendi dosyasını bırakır. İki kişi aynı anda gönderse bile çakışma önlenmiş
   değil, *yapısal olarak imkânsızdır*. Ayrı bir “gelen kutusu” yoktur: dosya
   baştan itibaren kalıcı yerindedir. Personel bu klasöre yazabilir ama içini
   göremez; `kaizen\` altında gördüğü tek şey çalışma kitabı ve `yonetim\`
   klasörünün adıdır.
2. **Yazma hiçbir şeyin üzerine yazmaz, yarım kayıt da okunmaz.** Dosya
   doğrudan son adıyla, “varsa oluşturma” kipinde açılır: aynı adda bir dosya
   varsa işlem başarısız olur ve gönderim yeni bir numarayla yeniden denenir.
   Kaydın sonuna bir `kayit_sonu` satırı yazılır; yazma yarıda kesilirse bu
   satır oluşmaz ve okuyucu dosyayı yok sayar.
3. **Geçmiş silinmez.** Her değerlendirme yeni bir olay dosyasıdır. Bir
   önerinin güncel durumu, gönderim dosyası üzerine olayların zaman sırasıyla
   uygulanmasıyla bulunur. Silinebilecek bir "geçmiş alanı" yoktur.

Bunun sonucu: **çalışma kitapları tek doğruluk kaynağı değildir.** Ekrandaki
her şey klasörlerden yeniden üretilebilir. Kitap silinse, bozulsa veya yeniden
kurulsa veri kaybolmaz.

---

## Kaizen modeli

**Durum akışı (PDCA)**

| Durum | PDCA | Anlamı |
|---|---|---|
| Yeni | — | Henüz incelenmedi |
| Değerlendirmede | Planla | Kaizen ekibi inceliyor |
| Planlandı | Planla | Kabul edildi, uygulama planlanıyor |
| Pilot Uygulamada | Uygula | Sınırlı alanda deneniyor |
| Ölçümleniyor | Kontrol | Pilot sonuçları ölçülüyor |
| Standartlaştırıldı | Önlem | Yaygınlaştırıldı, süreç oldu |
| Beklemede / Reddedildi | — | Şartlar oluşmadı / gerekçesiyle kapatıldı |

**Öncelik matrisi** — etki ve efor 1–5 arası puanlanır:

| | Düşük efor (1–2) | Yüksek efor (3–5) |
|---|---|---|
| **Yüksek etki (3–5)** | Hızlı Kazanım — önce bunlar | Büyük Proje — planlama ve kaynak ister |
| **Düşük etki (1–2)** | Doldurma İşi — boş kapasiteyle | Değerlendirme Dışı — bu haliyle önerilmez |

**Tasarruf kuralı:** Yıllık saat ve TL toplamlarına yalnızca fayda gerçekleşmiş
sayılan durumlardaki öneriler girer (Pilot Uygulamada, Ölçümleniyor,
Standartlaştırıldı). Henüz uygulanmamış bir öneri tasarruf olarak sayılsaydı
pano, gerçekte olmayan bir kazancı yönetime raporlardı.

---

## Kaynak düzeni

```
kur.py                    iki kitabı sıfırdan üreten betik
kaynak\
  tasarim.py              tasarım sistemi (renk, tipografi, ölçü)
  uret_ortak.py           masthead, kart, form alanı, tablo desenleri
  uret_oneri.py           KaizenOneri.xlsm sayfa düzeni
  uret_yonetim.py         KaizenYonetim.xlsm sayfa düzeni + grafikler
  com_kurulum.py          VBA enjeksiyonu, düğmeler, koruma, .xlsm kaydı
  vba\                    10 modül + iki ThisWorkbook dosyası
testler\
  tum_testler.py          hepsini sırayla çalıştırır
  test_uretim.py          Excel açmadan yapısal ve tutarlılık kontrolleri
  test_uctan_uca.py       gerçek Excel'de gerçek makrolar
  test_eszamanlilik.py    paralel süreçlerle çakışma denemesi
  test_izinler.py         gerçek NTFS izinleriyle bırakma kutusu doğrulaması
  goruntu_al.py           ekranları PDF olarak dışa aktarır (tasarım incelemesi)
cikti\                    üretilen .xlsm dosyaları
```

Çalışma kitapları elle hazırlanmaz. `python kur.py` sayfaları, biçimlendirmeyi,
doğrulama listelerini, düğmeleri ve grafikleri kurar; VBA modüllerini içeri
aktarır. Kaynak kod düz metin dosyalarında durduğu için sürüm takibi
yapılabilir, karşılaştırılabilir, gözden geçirilebilir.

---

## Teknik notlar

Aşağıdakiler tasarımı belirleyen, kolayca gözden kaçan noktalardır:

- **UTF-8.** Tüm dosya okuma/yazma `ADODB.Stream` ile yapılır. VBA'nın
  yerleşik `Open/Print` komutu ANSI yazar ve Türkçe karakterler başka bir
  makinede bozulur.
- **VBA kaynak kodu.** Modüller `VBComponents.Import` ile değil,
  `CodeModule.AddFromString` ile aktarılır. Import dosyayı sistemin ANSI kod
  sayfasıyla okur; AddFromString Unicode üzerinden geçer ve kod sayfasından
  bağımsızdır.
- **Zaman damgası çözünürlüğü.** Olay dosyalarının adı 10 ms çözünürlüklü
  zaman damgası taşır. Saniye çözünürlüğü yetmezdi: aynı saniyede yazılan iki
  olayın sırasını rastgele ek belirler ve durum yanlış türetilebilirdi.
- **Bırakma kutusu ve yazma yöntemi.** `yonetim\oneriler\` klasöründe personelin
  listeleme, okuma ve **silme** yetkisi yoktur. Bu yüzden `.tmp` yazıp
  yeniden adlandırma yöntemi kullanılamaz: yeniden adlandırma silme yetkisi
  ister ve işletim sistemi reddeder. Dosya doğrudan son adıyla, “varsa
  oluşturma” kipinde yazılır.
- **Kısa öneri numarası.** `ON-YYMMDD-XXX` biçimindedir; rastgele ekin
  alfabesinden `0/O` ve `1/I` çıkarılmıştır, telefonda söylenirken
  karıştırılmasın diye. Numara kısaldığı için aynı numaranın iki kez üretilme
  ihtimali doğar; bu yüzden yazma işlemi hiçbir dosyanın üzerine yazmaz ve
  çakışmada yeni bir numarayla yeniden denenir.
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
2. **Kişi bazlı yetki yok.** Yönetim erişimi tek ortak şifredir. Kayıtlarda
   hangi Windows kullanıcısının işlem yaptığı yazar, ama şifre paylaşıldığı
   için bu kimlik doğrulama değil, yalnızca iz kaydıdır.
3. **Anlık bildirim yok.** Yeni öneri geldiğinde kimseye haber gitmez; Kaizen
   ekibi kitabı açıp *Önerileri Yenile* demelidir.
4. **OneDrive ile eşlenmiş klasör.** Geliştirme ve testler OneDrive dışında
   yapılmalıdır; üretim hedefi bir UNC paylaşımıdır (`KURULUM.md` madde 1).

---

## Başlarken

```
pip install openpyxl pywin32
python kur.py
python testler\tum_testler.py
```

Ardından `KURULUM.md`.
