Option Explicit

' ============================================================================
'  KaizenYonetim.xlsm -- ThisWorkbook
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


' Konsolda bir satira cift tiklamak o oneriyi degerlendirme ekraninda acar.
' Dugmeye gitmeden calisan bu kisayol, gunluk kullanimda en cok tekrarlanan
' islemi tek harekete indirir.
' NOT: Cancel ByRef olmak ZORUNDADIR (varsayilan). ByVal yazilirsa imza olayin
' tanimina uymaz ve bu modul DERLENMEZ; hata ancak ilk olay tetiklendiginde
' ortaya cikar. DerlemeSinamasi() bunu her testte yakalar.
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, _
                                            Cancel As Boolean)
    Dim no As String

    If StrComp(Sh.Name, modUI.SAYFA_KONSOL, vbTextCompare) <> 0 Then Exit Sub
    If Target.Row < modKonsolide.KONSOL_ILK_SATIR Then Exit Sub

    no = Trim$(CStr(Sh.Cells(Target.Row, modKonsolide.KONSOL_ILK_SUTUN).Value & ""))
    If Len(no) = 0 Then Exit Sub

    Cancel = True
    modDegerlendirme.OneriyiAc no
End Sub


' Etki ya da efor degistiginde oncelik sinifi aninda guncellenir; ekip
' puanlamanin sonucunu kaydetmeden gorur.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If StrComp(Sh.Name, modUI.SAYFA_DEGERLENDIRME, vbTextCompare) <> 0 Then Exit Sub

    On Error Resume Next
    If Not Application.Intersect(Target, Sh.Range("dg_etki")) Is Nothing _
       Or Not Application.Intersect(Target, Sh.Range("dg_efor")) Is Nothing Then
        Application.EnableEvents = False
        modDegerlendirme.OncelikGoster Sh        ' korumasini kendi yonetir
        Application.EnableEvents = True
    End If
    On Error GoTo 0
End Sub
