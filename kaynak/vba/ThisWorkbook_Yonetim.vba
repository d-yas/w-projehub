Option Explicit

' ============================================================================
'  ProjeYonetim.xlsm -- ThisWorkbook
'
'  BU KITAP VERI DEPOSUDUR. Butun oneriler ve butun degerlendirme gecmisi
'  icindeki iki cok gizli sayfada durur (bkz. modDepo). Bu yuzden acilistaki
'  uc is, ekranlardan once gelir:
'
'  1) YEDEK. Veri tek bir dosyada durdugu icin gunluk disk kopyasi tek gercek
'     sigortadir. Kilit hala bizdeyken alinir: o an dosyayi kimse
'     degistiremez, yani kopya tutarlidir.
'
'  2) SALT OKUNURA GEC. Ekip kitabi gun boyu acik tutar. Yazma kilidi bizde
'     kalirsa personel HICBIR oneri gonderemez. ChangeFileAccess kilidi
'     birakir; kitap ekranda calismaya devam eder, yalnizca kaydedilemez.
'
'  3) EKRANLARI KAPAT. Kitap her acilista sifre kapisina doner.
'
'  Sayfa korumasi UserInterfaceOnly:=True ile yeniden kurulur; bu bayrak
'  dosyada saklanmaz, aksi halde makrolar korumali sayfalara yazamaz.
' ============================================================================

Private Sub Workbook_Open()
    On Error Resume Next

    modDepo.YedekAl
    modDepo.SaltOkunuraGec

    modUI.KorumalariKur
    modUI.OturumKapat

    ' Giris ekranindaki bant, kurulumla ilgili iki sorunun gorulecegi TEK
    ' yerdir. Ikisi de sessizdir: kullanici ancak kaydetmeye calisinca fark
    ' eder ve o zaman da neden oldugunu anlamaz.
    '
    ' 1) YANLIS KONUM. OneDrive ile eslenen bir klasorde degerlendirme HIC
    '    kaydedilemez (bkz. modDepo.DepoOneDriveAltindaMi). Bunu, ekip bir
    '    degerlendirme yazmadan ONCE soylemek gerekir.
    ' 2) KILIT BIRAKILAMADI. O zaman personel oneri gonderemez.
    If modDepo.DepoOneDriveAltindaMi() Then
        modUI.BantYaz ThisWorkbook.Worksheets(modUI.SAYFA_GIRIS), "giris_bant", _
            "⚠  Bu kurulum OneDrive ile eşlenen bir klasörde. " & _
            "Değerlendirme KAYDEDİLEMEZ; kurulumu OneDrive dışına taşıyın.", _
            modTasarim.CLR_UYARI_ZEMIN, modTasarim.CLR_UYARI_YAZI
    ElseIf Not modDepo.SaltOkunurMu() Then
        modUI.BantYaz ThisWorkbook.Worksheets(modUI.SAYFA_GIRIS), "giris_bant", _
            "⚠  Bu kitap yazma kipinde açıldı. Kapatıp yeniden açın; " & _
            "aksi halde personel öneri gönderemez.", _
            modTasarim.CLR_UYARI_ZEMIN, modTasarim.CLR_UYARI_YAZI
    End If

    On Error GoTo 0
End Sub


' Kitabin kendisi ASLA kaydedilmez.
'
' Kitap salt okunur acilir, yani ekranda yapilan her sey yalnizca bellektedir.
' Bir kaydetme denemesi -- Ctrl+S ya da kapanistaki "kaydedilsin mi" -- iki
' turlu zarar verebilirdi: ya reddedilip kullaniciyi "Farkli Kaydet"e
' yonlendirir ve depo ikiye bolunurdu, ya da (kilit bir sekilde bizdeyse)
' BELLEKTEKI ESKI kopya diskin uzerine yazilir ve o arada personelin
' gonderdigi butun oneriler silinirdi.
'
' NOT: Cancel ByRef olmak ZORUNDADIR (varsayilan). ByVal yazilirsa imza olayin
' tanimina uymaz ve bu modul DERLENMEZ.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Cancel = True
    ThisWorkbook.Saved = True
    modUI.Bilgi "Bu kitap kaydedilmez; verinin tamamı zaten dosyanın " & _
                "içindedir ve öneriler ile değerlendirmeler kaydedildikleri " & _
                "anda diske yazılır." & vbCrLf & vbCrLf & _
                "Ekranda gördüğünüzü tazelemek için “Önerileri Yenile” " & _
                "düğmesine basın.", "Proje Öneri Yönetimi"
End Sub


' Kapanista "değişiklikler kaydedilsin mi?" sorusu cikmasin: kitap salt
' okunur ve bellekteki degisiklikler (acilan/gizlenen ekranlar, cekilen veri)
' bilerek atilir.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ThisWorkbook.Saved = True
    On Error GoTo 0
End Sub


' Bu modulun derlendigini kanitlar.
'
' ThisWorkbook yalnizca bir olay tetiklendiginde derlenir. Buradaki bir hata
' (ornegin yanlis olay imzasi) uretimde ve testlerde fark edilmez; sonra ilk
' cift tiklamada ya da ilk hucre degisikliginde kullaniciya ham bir Visual
' Basic penceresi olarak cikar. Testler bu fonksiyonu cagirarak modulu
' derlenmeye zorlar.
Public Function DerlemeSinamasi() As String
    DerlemeSinamasi = "ok"
End Function


' Listede bir satira cift tiklamak o oneriyi degerlendirme ekraninda acar.
' Dugmeye gitmeden calisan bu kisayol, gunluk kullanimda en cok tekrarlanan
' islemi tek harekete indirir.
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, _
                                            Cancel As Boolean)
    Dim no As String

    If StrComp(Sh.Name, modUI.SAYFA_LISTE, vbTextCompare) <> 0 Then Exit Sub
    If Target.Row < modKonsolide.LISTE_ILK_SATIR Then Exit Sub

    no = Trim$(CStr(Sh.Cells(Target.Row, modKonsolide.LISTE_ILK_SUTUN).Value & ""))
    If Len(no) = 0 Then Exit Sub

    Cancel = True
    modDegerlendirme.OneriyiAc no
End Sub
