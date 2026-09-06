Option Explicit

' ============================================================================
'  ProjeYonetim.xlsm -- ThisWorkbook
'
'  Kitap her acildiginda kapali konuma doner: yalnizca Giris ekrani gorunur.
'  Sayfa korumasi UserInterfaceOnly:=True ile yeniden kurulur; bu bayrak
'  dosyada saklanmaz, aksi halde makrolar korumali sayfalara yazamaz.
' ============================================================================

Private Sub Workbook_Open()
    On Error Resume Next
    modUI.KorumalariKur
    modUI.OturumKapat
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
' NOT: Cancel ByRef olmak ZORUNDADIR (varsayilan). ByVal yazilirsa imza olayin
' tanimina uymaz ve bu modul DERLENMEZ; hata ancak ilk olay tetiklendiginde
' ortaya cikar. DerlemeSinamasi() bunu her testte yakalar.
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

