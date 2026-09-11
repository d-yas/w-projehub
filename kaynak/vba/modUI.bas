Attribute VB_Name = "modUI"
Option Explicit

' ============================================================================
'  modUI -- ekran yonetimi, sifre kapisi, koruma, kurumsal mesajlar
'
'  Sayfa korumasi UserInterfaceOnly:=True ile kurulur: kullanici hucreleri
'  bozamaz ama VBA yazabilir. Bu bayrak kitap her acildiginda sifirlanir,
'  bu yuzden KorumalariKur her Workbook_Open'da yeniden cagrilir.
' ============================================================================

Public Const SAYFA_GIRIS As String = "Giriş"
Public Const SAYFA_FORM As String = "Öneri Formu"
Public Const SAYFA_LISTE As String = "Liste"
Public Const SAYFA_DEGERLENDIRME As String = "Değerlendirme"
Public Const SAYFA_PANO As String = "Pano"
Public Const SAYFA_LISTELER As String = "Listeler"
Public Const SAYFA_PANOVERI As String = "PanoVeri"
Public Const SAYFA_VERI As String = "Veri"
Public Const SAYFA_TAKIP As String = "Takip"


' ###########################################################################
'  SESSIZ MOD -- otomasyon icin
'
'  Bir MsgBox ya da InputBox, gorunmeyen bir Excel ornegini SONSUZA KADAR
'  kilitler: diyalog ekranda degildir ama cagrilan makro geri donmez. Testler
'  bu yuzden once sessiz modu acar; mesajlar ekrana cikmak yerine kaydedilir
'  ve SonMesaj() ile okunabilir. Boylece testler hem kilitlenmez hem de
'  kullaniciya ne soylendigini dogrulayabilir.
'
'  Sessiz mod yalnizca kod tarafindan acilir; normal kullanimda kapalidir.
' ###########################################################################

Private m_sessiz As Boolean
Private m_sonMesaj As String

Public Sub SessizModAyarla(ByVal acik As Boolean)
    m_sessiz = acik
    m_sonMesaj = ""
End Sub

Public Function SessizMi() As Boolean
    SessizMi = m_sessiz
End Function

' Son kullanici mesajini dondurur ve temizler.
Public Function SonMesaj() As String
    SonMesaj = m_sonMesaj
    m_sonMesaj = ""
End Function

Private Sub MesajKaydet(ByVal tur As String, ByVal metin As String)
    m_sonMesaj = tur & ": " & metin
End Sub


' ###########################################################################
'  SIFRE KAPISI
' ###########################################################################

' Sifreyi sorar. Dogruysa True doner. Iptal edilirse sessizce False doner.
' NOT: Bu sifre gercek bir guvenlik siniri degildir (bkz. modAyar).
Public Function SifreDogrula(ByVal beklenen As String, ByVal baslik As String) As Boolean
    Dim girilen As String

    If m_sessiz Then
        SifreDogrula = True
        Exit Function
    End If

    girilen = InputBox("Devam etmek için şifrenizi girin:", baslik)
    If Len(girilen) = 0 Then
        SifreDogrula = False                    ' iptal ya da bos
        Exit Function
    End If

    If StrComp(girilen, beklenen, vbBinaryCompare) = 0 Then
        SifreDogrula = True
    Else
        SifreDogrula = False
        Hata "Şifre hatalı.", baslik
    End If
End Function


' ###########################################################################
'  SAYFA GORUNURLUGU
' ###########################################################################

Public Function SayfaVarMi(ByVal ad As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    SayfaVarMi = Not (ws Is Nothing)
    On Error GoTo 0
End Function

Public Sub SayfaGoster(ByVal ad As String, Optional ByVal etkinlestir As Boolean = True)
    If Not SayfaVarMi(ad) Then Exit Sub
    ThisWorkbook.Worksheets(ad).Visible = xlSheetVisible
    If etkinlestir Then
        ' Etkinlestirme ve secim yalnizca gorsel rahatliktir. Penceresi
        ' olmayan bir Excel'de (otomasyon, gizli oturum) hata verir; bu hata
        ' isin kendisini durdurmamalidir.
        On Error Resume Next
        ThisWorkbook.Worksheets(ad).Activate
        ThisWorkbook.Worksheets(ad).Range("A1").Select
        On Error GoTo 0
        YenidenCiz
    End If
End Sub

Public Sub SayfaGizle(ByVal ad As String)
    If Not SayfaVarMi(ad) Then Exit Sub
    ' Son gorunur sayfa gizlenemez; once baska bir sayfa gorunur olmali.
    If GorunurSayfaSayisi() <= 1 Then Exit Sub
    ThisWorkbook.Worksheets(ad).Visible = xlSheetVeryHidden
End Sub

Private Function GorunurSayfaSayisi() As Long
    Dim ws As Object, n As Long
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then n = n + 1
    Next ws
    GorunurSayfaSayisi = n
End Function

' Oturum acildiginda calisma sayfalarini gosterir, giris ekranini gizler.
Public Sub OturumAc(ByVal sayfalar As Variant, ByVal ilkSayfa As String)
    Dim v As Variant
    Application.ScreenUpdating = False
    For Each v In sayfalar
        If SayfaVarMi(CStr(v)) Then ThisWorkbook.Worksheets(CStr(v)).Visible = xlSheetVisible
    Next v
    SayfaGoster ilkSayfa
    SayfaGizle SAYFA_GIRIS
    Application.ScreenUpdating = True
End Sub

' Kitabi kapali konuma dondurur: yalnizca giris ekrani gorunur kalir.
' Kaydetmeden once cagrilir; boylece kitap bir dahaki acilista sifre sorar.
Public Sub OturumKapat()
    Dim ws As Object

    If Not SayfaVarMi(SAYFA_GIRIS) Then Exit Sub

    Application.ScreenUpdating = False
    ThisWorkbook.Worksheets(SAYFA_GIRIS).Visible = xlSheetVisible
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, SAYFA_GIRIS, vbTextCompare) <> 0 Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws
    On Error Resume Next
    ThisWorkbook.Worksheets(SAYFA_GIRIS).Activate
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub


' ###########################################################################
'  KORUMA
' ###########################################################################

' Butun sayfalari UserInterfaceOnly ile korur. Kilitsiz birakilmis hucreler
' (form giris alanlari) yazilabilir kalir.
Public Sub KorumalariKur()
    Dim ws As Object
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Protect Password:=modAyar.SIFRE_KORUMA, _
                   UserInterfaceOnly:=True, _
                   DrawingObjects:=False, _
                   Contents:=True, _
                   Scenarios:=False, _
                   AllowFiltering:=True, _
                   AllowSorting:=True
        ws.EnableSelection = xlNoRestrictions
    Next ws
    On Error GoTo 0
End Sub

Public Sub KorumaKapa(ByVal ws As Object)
    On Error Resume Next
    ws.Unprotect Password:=modAyar.SIFRE_KORUMA
    On Error GoTo 0
End Sub

Public Sub KorumaAc(ByVal ws As Object)
    On Error Resume Next
    ws.Protect Password:=modAyar.SIFRE_KORUMA, _
               UserInterfaceOnly:=True, _
               DrawingObjects:=False, _
               Contents:=True, _
               Scenarios:=False, _
               AllowFiltering:=True, _
               AllowSorting:=True
    On Error GoTo 0
End Sub


' ###########################################################################
'  MESAJLAR
' ###########################################################################

Public Sub Bilgi(ByVal metin As String, Optional ByVal baslik As String = "Proje Öneri")
    If m_sessiz Then
        MesajKaydet "bilgi", metin
        Exit Sub
    End If
    MsgBox metin, vbInformation Or vbOKOnly, baslik
End Sub

Public Sub Hata(ByVal metin As String, Optional ByVal baslik As String = "Proje Öneri")
    If m_sessiz Then
        MesajKaydet "hata", metin
        Exit Sub
    End If
    MsgBox metin, vbExclamation Or vbOKOnly, baslik
End Sub

Public Function Onay(ByVal metin As String, Optional ByVal baslik As String = "Proje Öneri") As Boolean
    If m_sessiz Then
        MesajKaydet "onay", metin
        Onay = True
        Exit Function
    End If
    Onay = (MsgBox(metin, vbQuestion Or vbYesNo Or vbDefaultButton2, baslik) = vbYes)
End Function


' ###########################################################################
'  EKRAN YARDIMCILARI
' ###########################################################################

' Bu iki yordam yalnizca hiz ve gorunum icindir; hicbiri isin sonucunu
' degistirmez. HizliModKapa cogu zaman bir HATA ISLEYICISI icinden cagrilir ve
' VBA'da isleyicinin kendi icinde olusan hata YAKALANAMAZ: makro cokup
' kullaniciya ham bir "Visual Basic" penceresi acar (otomasyonda ise gorunmeden
' kilitlenir). Bu yuzden ikisi de hataya tamamen dayanikli tutulur.
Public Sub HizliModAc()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.Cursor = xlWait
    On Error GoTo 0
End Sub

Public Sub HizliModKapa()
    On Error Resume Next
    Application.Cursor = xlDefault
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
'  YenidenCiz -- etkin pencereyi bastan boyatir.
'
'  Liste, kullanici BASKA bir ekrandayken ve ScreenUpdating kapaliyken
'  yeniden yazilir (modKonsolide.ListeyiCiz, modDegerlendirme.DegerlendirmeKaydet
'  icinden). Excel o sirada sayfanin ekran goruntusunu onbellekte tutar ve
'  sayfaya donuldugunde yalnizca gecersiz saydigi bolgeyi boyar.
'
'  Liste'de bolmeler dondurulmustur (C9). SayfaGoster donuste A1'i secer; A1
'  DONDURULMUS bolmededir, veri satirlari ise digerinde. Veri bolmesi gecersiz
'  sayilmazsa ESKI satirlarin yazisi ekranda kalir ve yeni satirlar ustune
'  cizilir -- satirlar ic ice gecmis gorunur.
'
'  Her bolmeyi bir satir kaydirip geri almak o bolmeyi gecersiz kilar. Gorunum
'  degismez: kaydirma konumu okunup aynen geri yazilir.
'
'  ScreenUpdating cagiranin biraktigi degere DONDURULUR, kosulsuz acilmaz:
'  OturumAc bu yordami ekran kapaliyken cagirir ve orada erken acmak sayfa
'  gizleme/gosterme trafigini kullaniciya seyrettirirdi.
'
'  Yalnizca gorunum icindir; penceresi olmayan bir oturumda (otomasyon,
'  testler) hata verebilir ve bu hata isin kendisini durdurmamalidir.
' ---------------------------------------------------------------------------
Public Sub YenidenCiz()
    Dim bolme As Object, satir As Long, eskiEkran As Boolean

    On Error Resume Next
    eskiEkran = Application.ScreenUpdating
    Application.ScreenUpdating = False
    For Each bolme In ActiveWindow.Panes
        satir = bolme.ScrollRow
        bolme.ScrollRow = satir + 1
        bolme.ScrollRow = satir
    Next bolme
    Application.ScreenUpdating = eskiEkran
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
'  Alan -- adlandirilmis bir aralik icin ustunde islem yapilabilir Range.
'
'  Ekran alanlarinin cogu birlesik hucredir (C10:F10 gibi) ve adlandirilmis
'  aralik yalnizca SOL UST hucreyi gosterir. Excel, birlesik bir hucrenin
'  PARCASI uzerinde ClearContents gibi islemleri reddeder ("bu işlemi
'  birleştirilmiş bir hücrede yapamayız"). MergeArea butun blogu dondurur;
'  boylece islem gecerli olur. Birlesik olmayan hucrelerde MergeArea hucrenin
'  kendisidir, yani bu yol her durumda dogrudur.
' ---------------------------------------------------------------------------
Public Function Alan(ByVal ws As Object, ByVal ad As String) As Range
    Set Alan = ws.Range(ad).MergeArea
End Function

' Bir hucreye kisa sureli durum bandi yazar (form ustundeki onay/uyari seridi).
Public Sub BantYaz(ByVal ws As Object, ByVal adres As String, ByVal metin As String, _
                   ByVal zeminHex As String, ByVal yaziHex As String)
    Dim r As Range
    KorumaKapa ws
    Set r = Alan(ws, adres)
    r.Cells(1, 1).Value = metin
    r.Interior.Color = modTasarim.HexRGB(zeminHex)
    r.Font.Color = modTasarim.HexRGB(yaziHex)
    r.Font.Name = modTasarim.FONT_AILE
    r.Font.Size = 10.5
    r.Font.Bold = True
    KorumaAc ws
End Sub

Public Sub BantTemizle(ByVal ws As Object, ByVal adres As String)
    Dim r As Range
    KorumaKapa ws
    Set r = Alan(ws, adres)
    r.ClearContents
    r.Interior.Color = modTasarim.HexRGB(modTasarim.CLR_BUZ_ZEMIN)
    KorumaAc ws
End Sub
