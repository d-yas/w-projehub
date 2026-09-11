Option Explicit

' ============================================================================
'  ProjeTakip.xlsm -- ThisWorkbook
'
'  Kitap her acildiginda BOS ekranla acilir. Onceki bir sorgunun sicil
'  numarasi ve sonucu ekranda kalmamalidir: kitap ortak bir bilgisayarda
'  acik birakilmis ya da sonucuyla birlikte bir kopyasi kaydedilmis olabilir.
'
'  Bu kitapta veri YOKTUR; sorgu depoyu her seferinde diskten okur
'  (bkz. modTakip). Kaydedilecek bir sey olmadigi icin kapanista soru sorulmaz.
'
'  Sayfa korumasi UserInterfaceOnly:=True ile yeniden kurulur; bu bayrak
'  dosyada saklanmaz, aksi halde makrolar korumali sayfaya yazamaz.
' ============================================================================

Private Sub Workbook_Open()
    On Error Resume Next
    modUI.KorumalariKur
    modTakip.EkraniSifirla
    On Error GoTo 0
End Sub


' Kapanista "değişiklikler kaydedilsin mi?" sorusu cikmasin: ekrandaki sorgu
' sonucu gecicidir ve bilerek atilir.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    ThisWorkbook.Saved = True
    On Error GoTo 0
End Sub


' Bu modulun derlendigini kanitlar. ThisWorkbook yalnizca bir olay
' tetiklendiginde derlenir; buradaki bir hata aksi halde ancak kullanicinin
' karsisinda ortaya cikar. Testler bu fonksiyonu cagirir.
Public Function DerlemeSinamasi() As String
    DerlemeSinamasi = "ok"
End Function
