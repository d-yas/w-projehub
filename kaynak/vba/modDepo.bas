Attribute VB_Name = "modDepo"
Option Explicit

' NOT: Modul duzeyi bildirimler (Const, Dim, Declare) VBA'da YALNIZCA burada,
' ilk yordamdan once bulunabilir. Yordamlarin arasina konan bir bildirim
' modulu DERLETMEZ ve VBA yordamlari tek tek derledigi icin hata ancak o
' yordam ilk cagrildiginda ortaya cikar (test_uretim.py bunu denetler).

' ============================================================================
'  modDepo -- VERI DEPOSU: yonetim kitabinin kendisi
'
'  Sistemde ARTIK KAYIT DOSYASI YOKTUR. Ne "oneriler\<yil>\PRJ-....txt" ne de
'  "degerlendirme\...txt". Butun veri yonetim kitabinin icindeki iki cok gizli
'  sayfada durur:
'
'      Oneriler  -- satir basina bir gonderim. DEGISMEZDIR, guncellenmez.
'      Olaylar   -- satir basina bir degerlendirme. YALNIZCA EKLENIR.
'
'  Bir onerinin guncel durumu, gonderim satirinin uzerine kendi olaylarinin
'  satir sirasiyla oynatilmasiyla bulunur (modKonsolide). Silinebilecek bir
'  "gecmis alani" yoktur; ekle-only guvencesi eskiden dosya adlarindan
'  geliyordu, artik satir sirasindan geliyor.
'
'  ---------------------------------------------------------------------
'  ESZAMANLILIK -- eski kuralin yerini alan yeni kural
'
'  Eski tasarimin kurali "kimse ortak bir dosyaya yazmaz" idi. Yeni kural:
'  HERKES AYNI DOSYAYA, KISA ve DISLAYICI bir islemle yazar:
'
'      kilidi al -> ac (yazma kipi) -> satir ekle -> Save -> Close -> kilidi birak
'                                                              (~3 saniye)
'
'  UC KATMAN VARDIR ve ucu de gereklidir:
'
'  1) KILIT DOSYASI (KilidiAl). Gercek karsilikli dislama. Excel'in kendi
'     kilidi kaydetme aninda kisa sureligine birakildigi icin TEK BASINA
'     YETMEZ; ayrintisi ve olculen kanit KilidiAl'in basindadir.
'
'  2) ACILAN KITABIN ReadOnly OZELLIGI. Dosya baskasinda aciksa Excel HATA
'     VERMEZ, dosyayi SESSIZCE SALT OKUNUR acar. Catisma sinyali hata degil,
'     ReadOnly'dir; yalnizca hataya bakan bir kod yazdigini sanirdi.
'
'  3) SAVE'IN SONUCU (KaydetmeyiDene). Save hatasi yutulursa satir diske
'     inmedigi halde cagirana basari doner; sonraki yazici da ayni numarayi
'     uretir. Bu yuzden Save asla varsayilmaz.
'
'  Her katmanda catisma olursa geri cekilip yeniden denenir. Butce duvar
'  saatiyle olculur; bitince cagiran tarafa anlasilir bir hata gider ve form
'  icerigi kaybolmaz.
'
'  ---------------------------------------------------------------------
'  NEDEN AYRI, GIZLI BIR EXCEL ORNEGI
'
'  Depoya her erisim CreateObject("Excel.Application") ile acilan AYRI bir
'  Excel surecinden yapilir. Uc sebep:
'
'  1) Yonetim kitabi ekibin ekraninda zaten acik olabilir; ayni dosya ayni
'     Excel orneginde ikinci kez acilamaz.
'  2) Kullanicinin kendi Excel'inin ayarlarina (olaylar, uyarilar, ekran
'     guncelleme) hic dokunulmaz; kod bir hata alsa bile kullanicinin Excel'i
'     bozulmus ayarlarla kalmaz.
'  3) Depo kitabinin Workbook_Open'i AutomationSecurity ile tumden kapatilir.
'
'  Bedeli ornek basina ~1 saniyedir ve olculmustur.
'
'  ---------------------------------------------------------------------
'  BU MODUL KULLANICIYLA KONUSMAZ
'
'  Gorunmez bir Excel'de acilan HERHANGI bir diyalog makroyu sonsuza kadar
'  bekletir. Bu yuzden burada MsgBox/InputBox yoktur; hatalar Err.Raise ile
'  yukari cikar ve ekran modulleri onlari modUI uzerinden gosterir. Ayni
'  sebeple depo HER ZAMAN parolasiyla acilir: parolasiz bir Open, gorunmez
'  Excel'de parola diyalogu acar ve makro geri donmez (olculdu).
' ============================================================================

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' --- Depo sayfalari (uret_yonetim.py ile ayni olmali) --------------------
Public Const SAYFA_ONERILER As String = "Oneriler"
Public Const SAYFA_OLAYLAR As String = "Olaylar"

' --- "Oneriler" sutun duzeni --------------------------------------------
Public Const O_SEMA As Long = 1
Public Const O_ONERI_NO As Long = 2
Public Const O_TARIH As Long = 3
Public Const O_AD_SOYAD As Long = 4
Public Const O_SICIL_NO As Long = 5
Public Const O_MEVCUT_DURUM As Long = 6
Public Const O_ONERI_BASLIGI As Long = 7
Public Const O_COZUM_ONERISI As Long = 8
Public Const O_BEKLENEN_FAYDA As Long = 9
Public Const O_GONDEREN_BILGISAYAR As Long = 10
Public Const O_GONDEREN_KULLANICI As Long = 11
Public Const O_SUTUN_SAYISI As Long = 11

' --- "Olaylar" sutun duzeni ----------------------------------------------
Public Const E_SEMA As Long = 1
Public Const E_ONERI_NO As Long = 2
Public Const E_OLAY_TARIHI As Long = 3
Public Const E_YENI_DURUM As Long = 4
Public Const E_KARAR_NOTU As Long = 5
Public Const E_DEGERLENDIREN_KULLANICI As Long = 6
Public Const E_DEGERLENDIREN_BILGISAYAR As Long = 7
Public Const E_SUTUN_SAYISI As Long = 7

' --- Zaman butceleri ----------------------------------------------------
' Butceler duvar saatiyle olculur, deneme sayisiyla degil: bir denemenin
' maliyeti (ac + kaydet + kapat, ~3 sn) makineye ve aga gore degisir.
'
' KILIT butcesi genistir cunku burada SIRA BEKLENIR: onumuzde bekleyen her
' yazma yaklasik uc saniye surer. YAZMA butcesi darcadir cunku o noktada kilit
' bizdedir ve dosyanin hemen acilmasi beklenir.
Private Const AZAMI_KILIT_SN As Double = 60#
Private Const AZAMI_YAZMA_SN As Double = 20#
Private Const AZAMI_OKUMA_SN As Double = 30#
Private Const BEKLEME_MIN_MS As Long = 300
Private Const BEKLEME_MAX_MS As Long = 900

' Yedek klasorunde taranacak azami dosya sayisi -- bkz. EskiYedekleriSil.
Private Const AZAMI_YEDEK_TARAMA As Long = 500

' Ayni acilis icinde kac kez Save denenecegi -- bkz. KaydetmeyiDene.
Private Const AZAMI_KAYIT_DENEMESI As Long = 3

' --- Yazma kilidi (bkz. KilidiAl) ---------------------------------------
Private Const KILIT_UZANTI As String = ".kilit"

' Bu kadar saniyeden eski bir kilit, sahibi cokmus sayilip kaldirilir.
'
' Deger bes dakikadir ve MAKINE SAATLERININ YAKIN OLDUGUNU varsayar: kilidin
' yasi, onu olusturan makinenin yazdigi dosya zamani ile BIZIM saatimizin
' farkidir. Bir alan ortaminda saatler Kerberos yuzunden zaten bes dakikadan
' fazla ayrilamaz, o yuzden esik oraya konuldu. Yanlis yere "bayat" denmesinin
' bedeli dosyanin bozulmasi degil, kilidin bir kez atlanmasidir -- yani
' kilitten onceki eski riskin nadir bir tekrari.
Private Const KILIT_BAYAT_SN As Double = 300#

' --- Excel sabitleri (surumden bagimsiz olsun diye sayisal) --------------
Private Const XL_YUKARI As Long = -4162          ' xlUp
Private Const XL_SALT_OKUNUR As Long = 3         ' xlReadOnly
Private Const MSO_GUVENLIK_KAPALI As Long = 3    ' msoAutomationSecurityForceDisable

Private m_tohumAtildi As Boolean

' Kilit su anda BIZDE mi? KilidiBirak buna bakar.
'
' Gerekcesi ince ama gercek: OlayEkleVeCek kilidi yazma biter bitmez birakir,
' sonra salt okunur cekme yapar. O cekme sirasinda bir hata olursa hata
' isleyicisi yine KilidiBirak cagirir -- ve bu arada kilidi BASKASI almis
' olabilir. Bayrak olmadan o kisinin kilidini silerdik ve iki yazici ayni anda
' calisirdi; onlemek icin var oldugu durumun ta kendisi.
Private m_kilitBizde As Boolean


' ###########################################################################
'  KAYIT ALANLARI -- sutun sirasinin tek kaynagi
' ###########################################################################

' Sozluk anahtarlari, sutun sirasiyla. Basliklar uret_yonetim.py'de ayni
' sirayla yazilir; test_uretim.py ikisini karsilastirir.
Public Function OneriAlanlari() As Variant
    OneriAlanlari = Array("sema", "oneri_no", "tarih", _
                          "ad_soyad", "sicil_no", _
                          "mevcut_durum", _
                          "oneri_basligi", "cozum_onerisi", "beklenen_fayda", _
                          "gonderen_bilgisayar", "gonderen_kullanici")
End Function

Public Function OlayAlanlari() As Variant
    OlayAlanlari = Array("sema", "oneri_no", "olay_tarihi", _
                         "yeni_durum", "karar_notu", _
                         "degerlendiren_kullanici", "degerlendiren_bilgisayar")
End Function


' ###########################################################################
'  KURULUM KONUMU -- OneDrive ile eslenen klasorde YAZMA calismaz
'
'  Olculerek bulundu: OneDrive ile eslenen bir klasordeki kitap bir Excel
'  orneginde acik oldugu surece -- salt okunur bile olsa -- ikinci bir Excel
'  sureci onu yazma kipinde acamiyor. Excel hata vermiyor, sessizce salt
'  okunur aciyor; DepoAc bunu dogru sekilde catisma sayip yeniden deniyor ve
'  butce dolunca "baskasi kullaniyor" diyor. Oysa kimse kullanmiyor: konum
'  yanlis.
'
'  Ekip kitabi ekranda HER ZAMAN acik oldugu icin, boyle bir klasorde
'  degerlendirme HIC kaydedilemez. O yuzden yirmi saniye beklemek yerine
'  DURUM ADIYLA soylenir. Gonderim tarafi engellenmez: personelin Excel'inde
'  yonetim kitabi acik olmadigi icin gonderim orada da calisir.
' ###########################################################################

Public Function DepoOneDriveAltindaMi() As Boolean
    On Error Resume Next
    DepoOneDriveAltindaMi = modDosyaIO.OneDriveAltindaMi(modAyar.YonetimKitapYolu())
    Err.Clear
    On Error GoTo 0
End Function

Public Function OneDriveAciklamasi() As String
    OneDriveAciklamasi = _
        "Bu kurulum OneDrive ile eşlenen bir klasörde duruyor ve orada " & _
        "değerlendirme kaydedilemez." & vbCrLf & vbCrLf & _
        "OneDrive, yönetim kitabı ekranda açık olduğu sürece dosyayı başka " & _
        "hiçbir işleme yazdırmıyor." & vbCrLf & vbCrLf & _
        "Çözüm: kurulumu OneDrive dışına, tercihen ağ paylaşımına taşıyın " & _
        "(bkz. KURULUM.md madde 1)."
End Function


' ###########################################################################
'  DIS KAPI -- her biri kendi gizli Excel ornegini acar ve kapatir
' ###########################################################################

' Bir gonderimi depoya ekler ve uretilen oneri numarasini dondurur.
' KILIT ONCE, EXCEL SONRA. Sirasi onemlidir: bekleyen bir yazici, sirasini
' beklerken bosta duran bir Excel ornegi tutmamalidir. Dort taraf ayni anda
' gonderim yaptiginda bu, onlarca gereksiz Excel surecine ve makineyi
' tikayacak kadar yuke yol aciyordu.
Public Function OneriEkle(ByVal sozluk As Object) As String
    Dim app As Object, hataMetni As String, aciklama As String

    If Not KilidiAl(aciklama, AZAMI_KILIT_SN) Then
        Err.Raise vbObjectError + 949, "modDepo.OneriEkle", aciklama
    End If

    On Error GoTo Hata
    Set app = GizliExcelAc()
    OneriEkle = IcOneriEkle(app, sozluk)
    GizliExcelKapat app
    KilidiBirak
    Exit Function

Hata:
    ' Aciklama ILK is olarak alinir: sonraki her cagri Err'i temizler.
    hataMetni = Err.Description
    GizliExcelKapat app
    KilidiBirak
    Err.Raise vbObjectError + 941, "modDepo.OneriEkle", hataMetni
End Function

' Bir degerlendirme olayini depoya ekler; yazildigi satir numarasini dondurur.
Public Function OlayEkle(ByVal sozluk As Object) As Long
    Dim app As Object, hataMetni As String, aciklama As String

    ' Yirmi saniye bekleyip yaniltici bir mesaj vermek yerine, durumu adiyla
    ' soyle (bkz. DepoOneDriveAltindaMi).
    If DepoOneDriveAltindaMi() Then
        Err.Raise vbObjectError + 946, "modDepo.OlayEkle", OneDriveAciklamasi()
    End If

    If Not KilidiAl(aciklama, AZAMI_KILIT_SN) Then
        Err.Raise vbObjectError + 948, "modDepo.OlayEkle", aciklama
    End If

    On Error GoTo Hata
    Set app = GizliExcelAc()
    OlayEkle = IcOlayEkle(app, sozluk)
    GizliExcelKapat app
    KilidiBirak
    Exit Function

Hata:
    hataMetni = Err.Description
    GizliExcelKapat app
    KilidiBirak
    Err.Raise vbObjectError + 942, "modDepo.OlayEkle", hataMetni
End Function

' Olayi yazar ve ARDINDAN depoyu ceker -- tek Excel ornegiyle iki is.
' Degerlendirme kaydedildikten sonra ekranin tazelenmesi gerektigi icin
' ikisi bir arada yapilir; ayri ayri yapilsa iki ornek baslatilirdi.
Public Function OlayEkleVeCek(ByVal sozluk As Object, _
                              ByVal hedefKitap As Object) As Long
    Dim app As Object, hataMetni As String, aciklama As String

    If DepoOneDriveAltindaMi() Then
        Err.Raise vbObjectError + 946, "modDepo.OlayEkleVeCek", OneDriveAciklamasi()
    End If

    If Not KilidiAl(aciklama, AZAMI_KILIT_SN) Then
        Err.Raise vbObjectError + 947, "modDepo.OlayEkleVeCek", aciklama
    End If

    On Error GoTo Hata
    Set app = GizliExcelAc()
    OlayEkleVeCek = IcOlayEkle(app, sozluk)

    ' Kilit YAZMA bitince birakilir; sonraki cekme salt okunurdur ve kimseyi
    ' bekletmemelidir.
    KilidiBirak

    IcDepoyuCek app, hedefKitap
    GizliExcelKapat app
    Exit Function

Hata:
    hataMetni = Err.Description
    GizliExcelKapat app
    KilidiBirak
    Err.Raise vbObjectError + 943, "modDepo.OlayEkleVeCek", hataMetni
End Function

' Diskteki depoyu okuyup hedef kitabin kendi Oneriler/Olaylar sayfalarina
' kopyalar. Ekranlar HER ZAMAN kendi kitaplarindaki bu kopyadan okur.
Public Sub DepoyuCek(ByVal hedefKitap As Object)
    Dim app As Object, hataMetni As String

    Set app = GizliExcelAc()
    On Error GoTo Hata
    IcDepoyuCek app, hedefKitap
    GizliExcelKapat app
    Exit Sub

Hata:
    hataMetni = Err.Description
    GizliExcelKapat app
    Err.Raise vbObjectError + 944, "modDepo.DepoyuCek", hataMetni
End Sub


' ###########################################################################
'  GIZLI EXCEL ORNEGI
' ###########################################################################

Public Function GizliExcelAc() As Object
    Dim app As Object

    Set app = CreateObject("Excel.Application")

    ' Hepsi hataya dayanikli: bir ozellik kurulamazsa is durmamali, ama
    ' AutomationSecurity kurulamazsa depo kitabinin Workbook_Open'i calisir.
    ' O da zararsizdir (yedek alir, salt okunura gecmeye calisir), yalnizca
    ' yavaslatir.
    On Error Resume Next
    app.Visible = False
    app.DisplayAlerts = False
    app.EnableEvents = False
    app.ScreenUpdating = False
    app.AskToUpdateLinks = False
    app.AutomationSecurity = MSO_GUVENLIK_KAPALI
    On Error GoTo 0

    Set GizliExcelAc = app
End Function

' ---------------------------------------------------------------------------
'  GizliExcelKapat -- ornegi HER hata yolunda kapatir.
'
'  Kapatilmayan bir ornek arkada gorunmez bir EXCEL.EXE olarak kalir ve
'  actigi dosyanin kilidini tutmaya devam eder: bir sonraki gonderim otuz
'  saniye bekleyip "baskasi kullaniyor" der, oysa kimse kullanmiyordur.
'
'  Quit'in gercekten kapatmasi icin O ORNEGE AIT hicbir nesne referansi elde
'  kalmamalidir -- bir Worksheet ya da Range degiskeni bile yeter. Cagiran
'  yordamlar bu yuzden kendi sayfa degiskenlerini once Nothing yapar.
'
'  Dongu en fazla kitap sayisi kadar doner: kapanmayan bir kitapta sonsuza
'  kadar denemek, hata halinde makronun hic geri donmemesi demekti.
' ---------------------------------------------------------------------------
Public Sub GizliExcelKapat(ByRef app As Object)
    Dim wb As Object
    Dim kalan As Long

    On Error Resume Next
    If Not app Is Nothing Then
        kalan = app.Workbooks.Count + 2
        Do While app.Workbooks.Count > 0 And kalan > 0
            kalan = kalan - 1
            Err.Clear
            Set wb = app.Workbooks(1)
            wb.Saved = True
            wb.Close SaveChanges:=False
            Set wb = Nothing
            If Err.Number <> 0 Then Exit Do
        Loop
        Err.Clear
        app.Quit
    End If
    Set app = Nothing
    Err.Clear
    On Error GoTo 0
End Sub


' ###########################################################################
'  DEPOYU ACMA -- catismada geri cekil, yeniden dene
' ###########################################################################

' Depoyu acar. Basarisizsa Nothing doner ve aciklama doldurulur.
'
' saltOkunur=False istendiginde acilan kitap ReadOnly ise BASKASI YAZIYOR
' demektir: kitap kapatilir ve yeniden denenir. Bu kontrol sartir -- Excel
' bu durumda hata vermez, sessizce salt okunur acar.
Public Function DepoAc(ByVal app As Object, ByVal saltOkunur As Boolean, _
                       Optional ByRef aciklama As String, _
                       Optional ByVal butceSn As Double = -1#) As Object
    Dim wb As Object
    Dim yol As String
    Dim baslangic As Double

    aciklama = ""
    yol = modAyar.YonetimKitapYolu()

    If Not modDosyaIO.DosyaVarMi(yol) Then
        aciklama = "Yönetim kitabı bulunamadı:" & vbCrLf & yol
        Exit Function
    End If

    ' Butce cagiran taraftan gelebilir: bir yazma islemi bastan denenirse
    ' acma ve kaydetme denemelerinin TOPLAMI kullaniciyi bekletmemelidir.
    If butceSn < 0# Then butceSn = AZAMI_OKUMA_SN

    baslangic = Timer
    Do
        Set wb = Nothing

        On Error Resume Next
        Set wb = app.Workbooks.Open( _
                    Filename:=yol, _
                    UpdateLinks:=0, _
                    ReadOnly:=saltOkunur, _
                    Password:=modAyar.SIFRE_DOSYA, _
                    IgnoreReadOnlyRecommended:=True, _
                    Notify:=False, _
                    AddToMru:=False)
        If Err.Number <> 0 Then
            aciklama = "Yönetim kitabı açılamadı (hata " & Err.Number & "): " & _
                       Err.Description
            Err.Clear
            Set wb = Nothing
        End If
        On Error GoTo 0

        If Not wb Is Nothing Then
            If saltOkunur Then
                Set DepoAc = wb
                Exit Function
            End If

            If Not wb.ReadOnly Then
                Set DepoAc = wb
                Exit Function
            End If

            ' Yazma istendi ama salt okunur acildi: depo su anda baskasinda.
            aciklama = "Yönetim kitabı şu anda başka bir kullanıcı " & _
                       "tarafından kullanılıyor."
            KitapKapat wb, False
        End If

        GeriCekil
    Loop While GecenSaniye(baslangic) < butceSn

    If Len(aciklama) = 0 Then
        aciklama = "Yönetim kitabına " & CLng(butceSn) & _
                   " saniye içinde erişilemedi."
    End If
End Function

' ---------------------------------------------------------------------------
'  KaydetmeyiDene -- Save'in BASARILI olup olmadigini soyler.
'
'  BU KONTROL SISTEMIN EN KRITIK YERIDIR. Save bir kez sessizce yutuldugu icin
'  su gorulmustu: dokuz paralel gonderimden biri KAYBOLDU ve bir numara IKI
'  KEZ verildi. Sebep zincirleme: bir yazicinin Save'i basarisiz oldu, hata
'  "On Error Resume Next" altinda yutuldu, cagiran taraf numarayi BASARI gibi
'  dondurdu; satir diske hic inmedigi icin sonraki yazici ayni "en buyuk + 1"
'  degerini hesapladi.
'
'  Bu yuzden Save'in sonucu asla varsayilmaz. Ayni acilis icinde birkac kez
'  denenir (cakisma cogunlukla anliktir); yine olmazsa cagiran taraf butun
'  islemi bastan dener ve o da olmazsa kullanici HATA gorur -- yazdiklari
'  formda durur. Sessiz kayip her ikisinden de kotudur.
'
'  Basarisiz bir Save diskteki dosyayi bozmaz: Excel once gecici bir dosya
'  yazip aslinin yerine koyar, adim tamamlanmazsa eski dosya oldugu gibi kalir.
' ---------------------------------------------------------------------------
Private Function KaydetmeyiDene(ByVal wb As Object, ByRef hataMetni As String) As Boolean
    Dim deneme As Long

    hataMetni = ""

    For deneme = 1 To AZAMI_KAYIT_DENEMESI
        On Error Resume Next
        Err.Clear
        wb.Save
        If Err.Number = 0 Then
            On Error GoTo 0
            hataMetni = ""
            KaydetmeyiDene = True
            Exit Function
        End If
        hataMetni = "Kaydedilemedi (hata " & Err.Number & "): " & Err.Description
        Err.Clear
        On Error GoTo 0
        Sleep BEKLEME_MIN_MS
    Next deneme

    KaydetmeyiDene = False
End Function

' ---------------------------------------------------------------------------
'  KitapKapat -- kitabi kapatir ve kilidi birakir.
'
'  YAZMA YOLU BURADAN KAYDETMEZ. Kaydetmenin basarili olup olmadigi bilinmek
'  zorundadir ve bu yordam tamamen hataya dayanikli olmak zorundadir; ikisi bir
'  arada olamaz. Kayit KaydetmeyiDene ile ayrica yapilir, sonucu denetlenir ve
'  buraya her zaman kaydet:=False ile gelinir.
'
'  Hataya dayaniklilik sart: cogu zaman bir hata isleyicisi icinden cagrilir ve
'  VBA'da isleyicinin kendi icinde olusan hata YAKALANAMAZ -- makro coker.
' ---------------------------------------------------------------------------
Public Sub KitapKapat(ByRef wb As Object, ByVal kaydet As Boolean)
    On Error Resume Next
    If Not wb Is Nothing Then
        If kaydet Then wb.Save
        wb.Saved = True
        wb.Close SaveChanges:=False
    End If
    Set wb = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

' ###########################################################################
'  YAZMA KILIDI
'
'  NEDEN EXCEL'IN KENDI KILIDI YETMIYOR -- ölçülerek bulundu.
'
'  Excel bir kitabi yazma kipinde acinca dosyayi kilitler ve ikinci bir ornek
'  onu yalnizca salt okunur acabilir. Buraya kadar iyi. Ama Excel KAYDEDERKEN
'  dosyayi yerinde degistirmez: gecici bir dosya yazip aslinin YERINE koyar. O
'  yer degistirme aninda eski dosyanin kilidi birakilir ve yenisininki henuz
'  alinmamistir. Pencere milisaniyeler surer, ama dort taraf ayni anda
'  yaziyorsa yakalanir:
'
'      A: aç → en büyük numara = N → satır N+1 → KAYDET ─┐
'      B:                          aç (pencere!) ────────┴→ eski içeriği görür
'      B: en büyük numara = N → satır N+1 → KAYDET  → A'nın satırını EZER
'
'  Sonuc: bir satir kaybolur ve AYNI NUMARA iki oneriye verilir. Eszamanlilik
'  testi bunu tam olarak boyle yakaladi (9 gönderim → 8 numara, 28 satır).
'
'  Cozum, Excel'e guvenmeyen gercek bir karsilikli dislama: yazmaya baslamadan
'  once "varsa olustur ma" kipinde bir kilit dosyasi olusturulur.
'  FileSystemObject.CreateTextFile(..., Overwrite:=False) bu isi atomik yapar
'  ve ag paylasimlarinda da gecerlidir. Kilit sifir baytlik, gecici bir
'  KOORDINASYON dosyasidir -- veri tasimaz ve islem biter bitmez silinir.
'
'  Sahibi cokerse kilit ortada kalir; bu yuzden KILIT_BAYAT_SN'den eski bir
'  kilit kaldirilir. Sure, en uzun islemin (30 sn butce) rahatca ustundedir.
' ###########################################################################

Private Function KilitYolu() As String
    KilitYolu = modAyar.YonetimKitapYolu() & KILIT_UZANTI
End Function

' Kilidi alir. Alamazsa False doner ve aciklama doldurulur.
Private Function KilidiAl(ByRef aciklama As String, ByVal butceSn As Double) As Boolean
    Dim fso As Object, akis As Object
    Dim yol As String
    Dim baslangic As Double

    yol = KilitYolu()
    baslangic = Timer

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then
        On Error GoTo 0
        aciklama = "Yazma kilidi kurulamadı."
        Exit Function
    End If
    On Error GoTo 0

    Do
        On Error Resume Next
        Err.Clear
        Set akis = fso.CreateTextFile(yol, False)      ' Overwrite:=False
        If Err.Number = 0 Then
            ' Kim tuttugu, bayat kilit tesbitinde ve sorun ararken lazim.
            akis.WriteLine Environ$("USERNAME") & "@" & Environ$("COMPUTERNAME") & _
                           "  " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
            akis.Close
            Set akis = Nothing
            Err.Clear
            On Error GoTo 0
            m_kilitBizde = True
            KilidiAl = True
            Exit Function
        End If
        Err.Clear
        Set akis = Nothing
        On Error GoTo 0

        BayatKilidiTemizle yol
        GeriCekil
    Loop While GecenSaniye(baslangic) < butceSn

    aciklama = "Yönetim kitabı şu anda başka bir kullanıcı tarafından " & _
               "kullanılıyor."
End Function

' Yalnizca kilit BIZDEYSE birakir; iki kez cagrilmasi zararsizdir.
Private Sub KilidiBirak()
    If Not m_kilitBizde Then Exit Sub
    m_kilitBizde = False

    On Error Resume Next
    Kill KilitYolu()
    Err.Clear
    On Error GoTo 0
End Sub

' Sahibi cokmus bir kilit sistemi tumden durdurmamalidir.
Private Sub BayatKilidiTemizle(ByVal yol As String)
    Dim yas As Double

    On Error Resume Next
    If Not modDosyaIO.DosyaVarMi(yol) Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If

    yas = (Now - FileDateTime(yol)) * 86400#
    If Err.Number = 0 Then
        If yas > KILIT_BAYAT_SN Then Kill yol
    End If
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub GeriCekil()
    Dim ms As Long
    If Not m_tohumAtildi Then
        Randomize
        m_tohumAtildi = True
    End If
    ms = BEKLEME_MIN_MS + Int(Rnd() * (BEKLEME_MAX_MS - BEKLEME_MIN_MS + 1))
    DoEvents
    Sleep ms
    DoEvents
End Sub

' Timer gece yarisinda sifirlanir; fark negatifse bir gun eklenir.
Private Function GecenSaniye(ByVal baslangic As Double) As Double
    Dim s As Double
    s = Timer - baslangic
    If s < 0 Then s = s + 86400#
    GecenSaniye = s
End Function


' ###########################################################################
'  IC ISLEMLER
' ###########################################################################

' ---------------------------------------------------------------------------
'  IcOneriEkle -- bir gonderim satiri ekler.
'
'  ISLEM ATOMIKTIR: satir yalnizca Save BASARILI olursa diske iner. Kaydetme
'  basarisiz olursa kitap KAYDEDILMEDEN kapatilir (diskteki dosya bozulmaz) ve
'  islem BASTAN denenir -- yeni bir acilis, yeni bir numara. Boylece "numara
'  verildi ama satir diske inmedi" durumu olusamaz; o durum hem oneriyi
'  kaybettirirdi hem de sonraki gonderime ayni numarayi verdirirdi.
'
'  Kilit bu noktada CAGIRANDADIR (bkz. OneriEkle); burasi yalnizca acma ve
'  kaydetme denemelerini yapar ve butcesi AZAMI_YAZMA_SN'dir. Kilit bizdeyken
'  dosyanin hemen acilmasi beklenir, o yuzden butce darcadir.
' ---------------------------------------------------------------------------
Private Function IcOneriEkle(ByVal app As Object, ByVal sozluk As Object) As String
    Dim wb As Object, ws As Object
    Dim aciklama As String, hataMetni As String
    Dim no As String, satir As Long
    Dim baslangic As Double, kalan As Double

    ' KILIT CAGIRANDADIR (bkz. OneriEkle). Burada yalnizca yazma yapilir.
    baslangic = Timer

    Do
        kalan = AZAMI_YAZMA_SN - GecenSaniye(baslangic)
        If kalan <= 0# Then Exit Do

        Set wb = DepoAc(app, False, aciklama, kalan)
        If wb Is Nothing Then Exit Do

        On Error GoTo Hata
        Set ws = wb.Worksheets(SAYFA_ONERILER)
        modUI.KorumaKapa ws

        ' Numara KILIT ICINDE uretilir: bu dosyaya o anda yalnizca biz
        ' yaziyoruz, dolayisiyla "en buyuk + 1" yarissiz ve sirali kalir.
        no = OneriNoUret(ws, Year(Now))
        sozluk("oneri_no") = no

        satir = SonSatir(ws) + 1
        SatirYaz ws, satir, SozluktenSatir(sozluk, OneriAlanlari())

        modUI.KorumaAc ws
        Set ws = Nothing             ' bkz. GizliExcelKapat: birakilmayan bir
        On Error GoTo 0              ' sayfa referansi Excel'i kapanmaz yapar

        If KaydetmeyiDene(wb, aciklama) Then
            KitapKapat wb, False     ' kaydedildi; kapatirken tekrar kaydetme
            IcOneriEkle = no
            Exit Function
        End If

        KitapKapat wb, False         ' satir diske inmedi; bastan denenecek
        GeriCekil
    Loop While GecenSaniye(baslangic) < AZAMI_YAZMA_SN

    If Len(aciklama) = 0 Then aciklama = "Öneri kaydedilemedi."
    Err.Raise vbObjectError + 950, "modDepo.IcOneriEkle", aciklama
    Exit Function

Hata:
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    Set ws = Nothing
    KitapKapat wb, False
    Err.Raise vbObjectError + 951, "modDepo.IcOneriEkle", hataMetni
End Function

' Bir degerlendirme olayi ekler. Atomiklik gerekcesi IcOneriEkle'dekiyle
' aynidir: kaydedilemeyen bir olay, kaydedilmis gibi gorunmemelidir.
Private Function IcOlayEkle(ByVal app As Object, ByVal sozluk As Object) As Long
    Dim wb As Object, ws As Object
    Dim aciklama As String, hataMetni As String
    Dim satir As Long
    Dim baslangic As Double, kalan As Double

    ' KILIT CAGIRANDADIR (bkz. OlayEkle / OlayEkleVeCek).
    baslangic = Timer

    Do
        kalan = AZAMI_YAZMA_SN - GecenSaniye(baslangic)
        If kalan <= 0# Then Exit Do

        Set wb = DepoAc(app, False, aciklama, kalan)
        If wb Is Nothing Then Exit Do

        On Error GoTo Hata
        Set ws = wb.Worksheets(SAYFA_OLAYLAR)
        modUI.KorumaKapa ws

        satir = SonSatir(ws) + 1
        SatirYaz ws, satir, SozluktenSatir(sozluk, OlayAlanlari())

        modUI.KorumaAc ws
        Set ws = Nothing
        On Error GoTo 0

        If KaydetmeyiDene(wb, aciklama) Then
            KitapKapat wb, False
            IcOlayEkle = satir
            Exit Function
        End If

        KitapKapat wb, False
        GeriCekil
    Loop While GecenSaniye(baslangic) < AZAMI_YAZMA_SN

    If Len(aciklama) = 0 Then aciklama = "Değerlendirme kaydedilemedi."
    Err.Raise vbObjectError + 952, "modDepo.IcOlayEkle", aciklama
    Exit Function

Hata:
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    Set ws = Nothing
    KitapKapat wb, False
    Err.Raise vbObjectError + 953, "modDepo.IcOlayEkle", hataMetni
End Function

Private Sub IcDepoyuCek(ByVal app As Object, ByVal hedefKitap As Object)
    Dim wb As Object
    Dim aciklama As String, hataMetni As String

    Set wb = DepoAc(app, True, aciklama)
    If wb Is Nothing Then
        Err.Raise vbObjectError + 954, "modDepo.IcDepoyuCek", aciklama
    End If

    On Error GoTo Hata
    SayfayiEsitle wb, hedefKitap, SAYFA_ONERILER, O_SUTUN_SAYISI
    SayfayiEsitle wb, hedefKitap, SAYFA_OLAYLAR, E_SUTUN_SAYISI
    KitapKapat wb, False
    Exit Sub

Hata:
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    KitapKapat wb, False
    Err.Raise vbObjectError + 955, "modDepo.IcDepoyuCek", hataMetni
End Sub

' Kaynak kitaptaki bir depo sayfasini hedef kitaptaki aynisina kopyalar.
Private Sub SayfayiEsitle(ByVal kaynakKitap As Object, ByVal hedefKitap As Object, _
                          ByVal sayfaAdi As String, ByVal sutunSayisi As Long)
    Dim kaynakWs As Object, hedefWs As Object
    Dim veri As Variant
    Dim n As Long, temizlenecek As Long

    Set kaynakWs = kaynakKitap.Worksheets(sayfaAdi)
    Set hedefWs = hedefKitap.Worksheets(sayfaAdi)

    veri = TabloOku(kaynakWs, sutunSayisi)

    modUI.KorumaKapa hedefWs
    temizlenecek = SonSatir(hedefWs)
    If temizlenecek >= 2 Then
        hedefWs.Range(hedefWs.Cells(2, 1), _
                      hedefWs.Cells(temizlenecek, sutunSayisi)).ClearContents
    End If

    If Not IsEmpty(veri) Then
        n = UBound(veri, 1)
        hedefWs.Cells(2, 1).Resize(n, sutunSayisi).NumberFormat = "@"
        hedefWs.Cells(2, 1).Resize(n, sutunSayisi).Value = veri
    End If
    modUI.KorumaAc hedefWs

    ' Kaynak sayfa GIZLI ORNEKTEDIR; referansi birakilmazsa o Excel kapanmaz.
    Set kaynakWs = Nothing
    Set hedefWs = Nothing
End Sub


' ###########################################################################
'  SATIR / TABLO YARDIMCILARI
' ###########################################################################

' Sozlugu, alan sirasina gore tek boyutlu bir dizeye cevirir.
' Cok satirli degerlerde satir sonlari vbLf'e indirgenir: Excel hucre ici
' satir sonu olarak onu kullanir, vbCrLf hucrede bos bir kare birakir.
Private Function SozluktenSatir(ByVal sozluk As Object, _
                                ByVal alanlar As Variant) As Variant
    Dim sonuc() As Variant
    Dim i As Long, deger As String

    ReDim sonuc(LBound(alanlar) To UBound(alanlar))
    For i = LBound(alanlar) To UBound(alanlar)
        deger = modDosyaIO.Al(sozluk, CStr(alanlar(i)))
        deger = Replace(deger, vbCrLf, vbLf)
        deger = Replace(deger, vbCr, vbLf)
        sonuc(i) = deger
    Next i

    SozluktenSatir = sonuc
End Function

' Bir satiri tek COM cagrisiyla yazar.
'
' Bicim ONCE metne cevrilir: aksi halde Excel "10045" sicil numarasini sayiya,
' "2026-09-08T10:11:51" tarihini tarihe cevirir ve basindaki sifirlari yer.
' Metin bicimi ayrica "=" ile baslayan bir metnin formul sanilmasini onler.
Private Sub SatirYaz(ByVal ws As Object, ByVal satir As Long, ByVal degerler As Variant)
    Dim hedef As Object
    Dim adet As Long

    adet = UBound(degerler) - LBound(degerler) + 1
    Set hedef = ws.Cells(satir, 1).Resize(1, adet)
    hedef.NumberFormat = "@"
    hedef.Value = degerler
End Sub

' Ilk sutuna gore son dolu satir. Basliklar 1. satirdadir, yani veri yoksa 1.
Public Function SonSatir(ByVal ws As Object) As Long
    Dim r As Long
    On Error Resume Next
    r = ws.Cells(ws.Rows.Count, 1).End(XL_YUKARI).Row
    If Err.Number <> 0 Then
        r = 1
        Err.Clear
    End If
    On Error GoTo 0
    If r < 1 Then r = 1
    SonSatir = r
End Function

' Bir depo sayfasinin veri satirlarini iki boyutlu dizi olarak dondurur.
' Veri yoksa Empty doner (cagiran taraf IsEmpty ile bakar).
Public Function TabloOku(ByVal ws As Object, ByVal sutunSayisi As Long) As Variant
    Dim son As Long

    son = SonSatir(ws)
    If son < 2 Then
        TabloOku = Empty
        Exit Function
    End If

    TabloOku = AralikDizisi(ws.Cells(2, 1).Resize(son - 1, sutunSayisi))
End Function

' ---------------------------------------------------------------------------
'  AralikDizisi -- Range.Value'yu HER ZAMAN iki boyutlu dizi olarak verir.
'
'  Excel tek hucrelik bir aralikta Value olarak dizi degil DEGERIN KENDISINI
'  dondurur. Cagiran taraf veri(i, 1) diye okumaya calisinca "Type mismatch"
'  alir; ustelik bu yalnizca depoda TAM BIR satir varken olur, yani ilk
'  gonderim calisir, ikincisi patlar. (Bu hata boyle bulundu.)
' ---------------------------------------------------------------------------
Public Function AralikDizisi(ByVal aralik As Object) As Variant
    Dim tek(1 To 1, 1 To 1) As Variant

    If aralik.Cells.Count = 1 Then
        tek(1, 1) = aralik.Value
        AralikDizisi = tek
    Else
        AralikDizisi = aralik.Value
    End If
End Function


' ###########################################################################
'  ONERI NUMARASI -- PRJ-2026-0001
' ###########################################################################

' Numara yil + dort haneli sirali sayidir.
'
' Sirali sayac eskiden IMKANSIZDI: ortak bir sayac dosyasi gerekiyordu ve
' "kimse ortak dosyaya yazmaz" kurali buna izin vermiyordu. Artik yazma
' zaten dislayici bir kilit altinda yapiliyor, yani sayaci okuyup bir
' artirmak yarissizdir. Numara okunur ve telefonda soylenebilir kaldi;
' ayrica kacinci onerinin geldigi dogrudan gorunur.
Public Function OneriNoUret(ByVal ws As Object, ByVal yil As Long) As String
    Dim onek As String, kuyruk As String
    Dim veri As Variant
    Dim son As Long, i As Long, n As Long, enBuyuk As Long
    Dim deger As String

    onek = modAyar.ONEK_ONERI_NO & "-" & Format$(yil, "0000") & "-"

    son = SonSatir(ws)
    If son >= 2 Then
        veri = AralikDizisi(ws.Cells(2, O_ONERI_NO).Resize(son - 1, 1))
        For i = LBound(veri, 1) To UBound(veri, 1)
            deger = Trim$(CStr(veri(i, 1) & ""))
            If Len(deger) > Len(onek) Then
                If StrComp(Left$(deger, Len(onek)), onek, vbTextCompare) = 0 Then
                    kuyruk = Mid$(deger, Len(onek) + 1)
                    If IsNumeric(kuyruk) Then
                        n = CLng(Val(kuyruk))
                        If n > enBuyuk Then enBuyuk = n
                    End If
                End If
            End If
        Next i
    End If

    OneriNoUret = onek & Format$(enBuyuk + 1, "0000")
End Function


' ###########################################################################
'  EKIP KITABI -- kilidi birakmak ve gunluk yedek
' ###########################################################################

' Kitabi SALT OKUNUR kipe alir ve dosya kilidini birakir.
'
' Bunsuz ekip kitabi gun boyu acik tuttugunda personel hicbir oneri
' gonderemezdi: depo baskasinda yazma kipinde acik olurdu. Cagri, kitap
' HENUZ KIRLENMEDEN yapilmalidir; Saved=True ve DisplayAlerts=False ile
' hicbir diyalog acilmadigi olculmustur.
Public Sub SaltOkunuraGec()
    Dim eskiUyari As Boolean

    On Error Resume Next
    If ThisWorkbook.ReadOnly Then Exit Sub

    eskiUyari = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True
    ThisWorkbook.ChangeFileAccess XL_SALT_OKUNUR
    Application.DisplayAlerts = eskiUyari
    Err.Clear
    On Error GoTo 0
End Sub

Public Function SaltOkunurMu() As Boolean
    On Error Resume Next
    SaltOkunurMu = ThisWorkbook.ReadOnly
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------------
'  YedekAl -- gunde bir disk kopyasi, son YEDEK_ADET kadari saklanir.
'
'  Veri artik TEK bir dosyada durdugu icin yedek bir sus degil, tek gercek
'  sigortadir: dosya silinirse ya da bozulursa butun oneriler ve butun
'  degerlendirme gecmisi gider.
'
'  SaveCopyAs DEGIL, disk kopyasi kullanilir: SaveCopyAs bellekteki durumu
'  yazar (ekranlarin acilmis/gizlenmis hali dahil), disk kopyasi ise dosyanin
'  bayt bayt aynisidir ve dolayisiyla parolasi da yerindedir.
'
'  Tamamen hataya dayaniklidir: yedek alinamamasi kitabin acilmasini
'  engellememelidir.
' ---------------------------------------------------------------------------
Public Sub YedekAl()
    Dim kaynak As String, klasor As String, hedef As String
    Dim fso As Object
    Dim deneme As Long

    On Error Resume Next

    kaynak = modAyar.YonetimKitapYolu()
    If Not modDosyaIO.DosyaVarMi(kaynak) Then Exit Sub

    klasor = modAyar.YedekKlasor()
    modDosyaIO.KlasorZinciriOlustur klasor
    If Not modDosyaIO.KlasorVarMi(klasor) Then Exit Sub

    hedef = klasor & "\" & YedekAdi(kaynak, Now)

    If Not modDosyaIO.DosyaVarMi(hedef) Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        For deneme = 1 To 3
            Err.Clear
            fso.CopyFile kaynak, hedef, True
            If Err.Number = 0 Then Exit For
            Sleep 250
        Next deneme
        Set fso = Nothing
    End If

    EskiYedekleriSil klasor, kaynak
    Err.Clear
    On Error GoTo 0
End Sub

' "ProjeYonetim.xlsm" + 8 Eylul 2026 -> "ProjeYonetim_20260908.xlsm"
' Ad kitabin GERCEK adindan turetilir; dosya yeniden adlandirilmis olabilir.
Public Function YedekAdi(ByVal kaynakYolu As String, ByVal gun As Date) As String
    Dim ad As String, govde As String, uzanti As String, p As Long

    ad = modDosyaIO.DosyaAdi(kaynakYolu)
    p = InStrRev(ad, ".")
    If p > 1 Then
        govde = Left$(ad, p - 1)
        uzanti = Mid$(ad, p)
    Else
        govde = ad
        uzanti = ".xlsm"
    End If

    YedekAdi = govde & "_" & Format$(gun, "yyyymmdd") & uzanti
End Function

' ---------------------------------------------------------------------------
'  EskiYedekleriSil -- en yeni YEDEK_ADET kopyayi birakir, gerisini siler.
'  Ad tarihle bittigi icin ada gore siralama = zamana gore siralama.
'
'  DIKKAT -- "On Error Resume Next" altinda Dir$ dongusu SONSUZA GIDEBILIR.
'  Dir$ hata verdiginde deger dondurmez; atama yapilmadigi icin degisken ESKI
'  degerinde kalir ve "Do While Len(ad) > 0" hep dogru olur. Klasor
'  listelenemedigi anda (yedek klasoru personele kapali oldugunda tam olarak
'  boyledir) makro geri donmez -- gorunmez bir Excel'de bu, kilitlenmenin ta
'  kendisidir. Olculerek bulundu; her Dir$ cagrisindan sonra Err denetlenir ve
'  ayrica bir ust sinir vardir.
' ---------------------------------------------------------------------------
Private Sub EskiYedekleriSil(ByVal klasor As String, ByVal kaynakYolu As String)
    Dim adlar() As String, sayi As Long, i As Long
    Dim ad As String, onek As String, govde As String, p As Long
    Dim desen As String

    On Error Resume Next

    ad = modDosyaIO.DosyaAdi(kaynakYolu)
    p = InStrRev(ad, ".")
    If p > 1 Then govde = Left$(ad, p - 1) Else govde = ad
    onek = govde & "_"
    desen = modDosyaIO.YolTemizle(klasor) & "\" & onek & "*"

    ReDim adlar(0 To 0)
    sayi = 0

    Err.Clear
    ad = Dir$(desen, vbNormal)
    If Err.Number <> 0 Then
        Err.Clear                       ' klasor listelenemiyor: yapacak is yok
        On Error GoTo 0
        Exit Sub
    End If

    Do While Len(ad) > 0 And sayi < AZAMI_YEDEK_TARAMA
        ReDim Preserve adlar(0 To sayi)
        adlar(sayi) = ad
        sayi = sayi + 1
        Err.Clear
        ad = Dir$
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
    Loop

    If sayi > modAyar.YEDEK_ADET Then
        AdlariSirala adlar, sayi
        For i = 0 To sayi - modAyar.YEDEK_ADET - 1
            Kill modDosyaIO.YolTemizle(klasor) & "\" & adlar(i)
        Next i
    End If

    Err.Clear
    On Error GoTo 0
End Sub

' Kucuk listeler icin ekleme siralamasi yeterlidir.
Private Sub AdlariSirala(ByRef liste() As String, ByVal sayi As Long)
    Dim i As Long, j As Long, gecici As String
    For i = 1 To sayi - 1
        gecici = liste(i)
        j = i - 1
        Do While j >= 0
            If StrComp(liste(j), gecici, vbTextCompare) > 0 Then
                liste(j + 1) = liste(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        liste(j + 1) = gecici
    Next i
End Sub


' ###########################################################################
'  TESTLER ICIN
' ###########################################################################

' Diskteki deponun satir sayilari: "oneri=N;olay=M".
' Hata durumunda "HATA: ..." doner; yakalanmamis bir VBA hatasi gorunmez bir
' Excel'de gorunmeyen bir pencere acar ve testi sonsuza kadar bekletir.
Public Function TestDepoSayilari() As String
    Dim app As Object, wb As Object
    Dim aciklama As String
    Dim oneri As Long, olay As Long

    On Error GoTo Hata

    Set app = GizliExcelAc()
    Set wb = DepoAc(app, True, aciklama)
    If wb Is Nothing Then
        GizliExcelKapat app
        TestDepoSayilari = "HATA: " & aciklama
        Exit Function
    End If

    oneri = SonSatir(wb.Worksheets(SAYFA_ONERILER)) - 1
    olay = SonSatir(wb.Worksheets(SAYFA_OLAYLAR)) - 1
    KitapKapat wb, False
    GizliExcelKapat app

    TestDepoSayilari = "oneri=" & oneri & ";olay=" & olay
    Exit Function

Hata:
    TestDepoSayilari = "HATA: " & Err.Number & " " & Err.Description
    KitapKapat wb, False
    GizliExcelKapat app
End Function
