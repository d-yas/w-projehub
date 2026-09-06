Option Explicit

' ============================================================================
'  ProjeOneri.xlsm -- ThisWorkbook
'
'  Kitap her acildiginda kapali konuma doner: yalnizca Giris ekrani gorunur.
'  Dosya nasil kaydedilmis olursa olsun sifre kapisi atlanamaz.
'
'  Sayfa korumasi UserInterfaceOnly:=True ile kurulur; bu bayrak dosyada
'  saklanmaz, her acilista yeniden verilmelidir. Aksi halde makrolar korumali
'  sayfaya yazamaz.
' ============================================================================

Private Sub Workbook_Open()
    On Error Resume Next
    modUI.KorumalariKur
    modUI.OturumKapat
    On Error GoTo 0
End Sub


' Bu modulun derlendigini kanitlar. ThisWorkbook yalnizca bir olay
' tetiklendiginde derlenir; buradaki bir hata aksi halde ancak kullanicinin
' karsisinda ortaya cikar. Testler bu fonksiyonu cagirir.
Public Function DerlemeSinamasi() As String
    DerlemeSinamasi = "ok"
End Function
