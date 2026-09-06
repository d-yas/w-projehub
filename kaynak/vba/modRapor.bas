Attribute VB_Name = "modRapor"
Option Explicit

' ============================================================================
'  modRapor -- yonetime giden parolali rapor
'
'  SISTEMDEKI TEK GERCEK GUVENLIK SINIRI BURASIDIR.
'  Ekran sifreleri VBA icinde duz metin oldugu icin gercek bir sinir degildir.
'  Rapor ise Excel'in kendi ECMA-376 (AES) sifrelemesiyle korunur; parolasiz
'  acilamaz.
'
'  Parola KODA GOMULMEZ. Rapor uretilirken kullaniciya sorulur ve hicbir yere
'  kaydedilmez -- boylece kitabi eline gecirmis biri raporu acamaz.
' ============================================================================

Private Const XL_XLSX As Long = 51             ' xlOpenXMLWorkbook

Private Const R_BASLIK_SATIR As Long = 8


' ###########################################################################
'  DUGME EYLEMLERI
' ###########################################################################

Public Sub RaporaGit()
    modUI.SayfaGoster modUI.SAYFA_RAPOR
End Sub

Public Sub RaporUret()
    Dim parola As String, parolaTekrar As String
    Dim donem As String, hedefDosya As String
    Dim veri As Variant, secili() As Long, adet As Long
    Dim ws As Object

    Set ws = RaporSayfasi()
    donem = Trim$(CStr(ws.Range("rpr_donem").Value & ""))
    If Len(donem) = 0 Then donem = "Tümü"

    veri = modKonsolide.VeriDizisi()
    If IsEmpty(veri) Then
        modUI.Hata "Raporlanacak öneri yok." & vbCrLf & vbCrLf & _
                   "Önce konsolda “Önerileri Yenile” düğmesine basın.", "Rapor"
        Exit Sub
    End If

    secili = DonemSuz(veri, donem, adet)
    If adet = 0 Then
        modUI.Hata "Seçilen dönemde (" & donem & ") öneri bulunamadı.", "Rapor"
        Exit Sub
    End If

    parola = InputBox( _
        "Rapor dosyası bu parolayla şifrelenecek." & vbCrLf & vbCrLf & _
        "Parola hiçbir yere kaydedilmez; kaybolursa rapor açılamaz." & vbCrLf & _
        "En az 6 karakter girin:", "Rapor parolası")
    If Len(parola) = 0 Then Exit Sub
    If Len(parola) < 6 Then
        modUI.Hata "Parola en az 6 karakter olmalıdır.", "Rapor"
        Exit Sub
    End If

    parolaTekrar = InputBox("Parolayı bir kez daha girin:", "Rapor parolası")
    If StrComp(parola, parolaTekrar, vbBinaryCompare) <> 0 Then
        modUI.Hata "Parolalar aynı değil. Rapor üretilmedi.", "Rapor"
        Exit Sub
    End If

    On Error GoTo Hata
    modUI.HizliModAc
    hedefDosya = KitabiUret(veri, secili, adet, donem, parola)
    modUI.HizliModKapa

    modUI.KorumaKapa ws
    ws.Range("rpr_son").Value = "Son üretilen rapor:  " & hedefDosya
    modUI.KorumaAc ws

    modUI.Bilgi "Rapor üretildi:" & vbCrLf & vbCrLf & hedefDosya & vbCrLf & vbCrLf & _
                adet & " öneri raporlandı. Dosya parolasız açılamaz.", "Rapor hazır"
    Exit Sub

Hata:
    modUI.HizliModKapa
    modUI.Hata "Rapor üretilemedi." & vbCrLf & vbCrLf & Err.Description, "Rapor"
End Sub


' ###########################################################################
'  DONEM SUZGECI
' ###########################################################################

Private Function DonemSuz(ByVal veri As Variant, ByVal donem As String, _
                          ByRef adet As Long) As Long()
    Dim secili() As Long, i As Long, n As Long
    Dim tarih As String, buYil As String, buAy As String

    n = UBound(veri, 1)
    ReDim secili(1 To n)
    buYil = Format$(Now, "yyyy")
    buAy = Format$(Now, "yyyy-mm")
    adet = 0

    For i = 1 To n
        tarih = CStr(veri(i, modKonsolide.V_TARIH))
        Select Case donem
            Case "Bu Yıl"
                If Left$(tarih, 4) = buYil Then
                    adet = adet + 1
                    secili(adet) = i
                End If
            Case "Bu Ay"
                If Left$(tarih, 7) = buAy Then
                    adet = adet + 1
                    secili(adet) = i
                End If
            Case Else                      ' "Tümü"
                adet = adet + 1
                secili(adet) = i
        End Select
    Next i

    DonemSuz = secili
End Function


' ###########################################################################
'  RAPOR KITABI
' ###########################################################################

Private Function KitabiUret(ByVal veri As Variant, ByRef secili() As Long, _
                            ByVal adet As Long, ByVal donem As String, _
                            ByVal parola As String) As String
    Dim rapor As Workbook
    Dim ozet As Object, tablo As Object
    Dim klasor As String, dosya As String
    Dim eskiUyari As Boolean

    klasor = modAyar.RaporKlasor()
    modDosyaIO.KlasorZinciriOlustur klasor
    dosya = klasor & "\KaizenRaporu_" & DonemEki(donem) & "_" & _
            Format$(Now, "yyyymmdd-hhnn") & ".xlsx"

    Set rapor = Application.Workbooks.Add
    Do While rapor.Worksheets.Count > 1
        rapor.Worksheets(rapor.Worksheets.Count).Delete
    Loop

    Set ozet = rapor.Worksheets(1)
    ozet.Name = "Özet"
    Set tablo = rapor.Worksheets.Add(After:=ozet)
    tablo.Name = "Öneriler"

    OzetYaz ozet, veri, secili, adet, donem
    TabloYaz tablo, veri, secili, adet

    ozet.Activate
    ozet.Range("A1").Select

    eskiUyari = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo Temizle
    If modDosyaIO.DosyaVarMi(dosya) Then Kill dosya
    ' Password:= Excel'in ECMA-376 AES sifrelemesini uygular.
    rapor.SaveAs Filename:=dosya, FileFormat:=XL_XLSX, Password:=parola
    rapor.Close SaveChanges:=False
    Application.DisplayAlerts = eskiUyari

    KitabiUret = dosya
    Exit Function

Temizle:
    On Error Resume Next
    rapor.Close SaveChanges:=False
    Application.DisplayAlerts = eskiUyari
    On Error GoTo 0
    Err.Raise vbObjectError + 940, "modRapor.KitabiUret", _
              "Rapor dosyası yazılamadı: " & dosya
End Function

Private Function DonemEki(ByVal donem As String) As String
    Select Case donem
        Case "Bu Yıl": DonemEki = Format$(Now, "yyyy")
        Case "Bu Ay":  DonemEki = Format$(Now, "yyyy-mm")
        Case Else:     DonemEki = "Tumu"
    End Select
End Function


' --------------------------------------------------------------------------
'  Ozet sayfasi -- kapak + gostergeler
' --------------------------------------------------------------------------
Private Sub OzetYaz(ByVal ws As Object, ByVal veri As Variant, _
                    ByRef secili() As Long, ByVal adet As Long, ByVal donem As String)
    Dim i As Long, k As Long, durum As String
    Dim bekleyen As Long, uygulanan As Long, kabul As Long, sonuclanan As Long
    Dim hizli As Long
    Dim saat As Double, tl As Double
    Dim satir As Long

    For k = 1 To adet
        i = secili(k)
        durum = CStr(veri(i, modKonsolide.V_DURUM))
        If modModel.BekliyorMu(durum) Then bekleyen = bekleyen + 1
        If modModel.KabulEdildiMi(durum) Then kabul = kabul + 1
        If modModel.SonuclandiMi(durum) Then sonuclanan = sonuclanan + 1
        If modModel.TasarrufSayilirMi(durum) Then
            uygulanan = uygulanan + 1
            saat = saat + Sayi(veri(i, modKonsolide.V_SAAT))
            tl = tl + Sayi(veri(i, modKonsolide.V_TL))
        End If
        If StrComp(CStr(veri(i, modKonsolide.V_ONCELIK)), _
                   modModel.ONCELIK_HIZLI, vbTextCompare) = 0 Then hizli = hizli + 1
    Next k

    ws.Cells.Interior.Color = modTasarim.HexRGB(modTasarim.CLR_BUZ_ZEMIN)
    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B").ColumnWidth = 38
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 2.5

    ' Kapak bandi
    With ws.Range("B2:C4")
        .Merge
        .Interior.Color = modTasarim.HexRGB(modTasarim.CLR_ANA_LACIVERT)
        .Font.Name = modTasarim.FONT_BASLIK_AILE
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = modTasarim.HexRGB(modTasarim.CLR_BEYAZ)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Value = "Kaizen Yönetim Raporu"
    End With
    ws.Rows("2:4").RowHeight = 24

    With ws.Range("B6:C6")
        .Merge
        .Value = "Dönem: " & donem & "     ·     Üretim: " & _
                 Format$(Now, "dd.mm.yyyy hh:nn") & "     ·     " & _
                 adet & " öneri"
        .Font.Name = modTasarim.FONT_AILE
        .Font.Size = 10
        .Font.Color = modTasarim.HexRGB(modTasarim.CLR_METIN_GRI)
    End With

    satir = R_BASLIK_SATIR
    BaslikYaz ws, satir, "Göstergeler"
    satir = satir + 1

    GostergeYaz ws, satir, "Toplam öneri", adet, "0": satir = satir + 1
    GostergeYaz ws, satir, "Değerlendirme bekleyen", bekleyen, "0": satir = satir + 1
    GostergeYaz ws, satir, "Uygulamaya geçmiş", uygulanan, "0": satir = satir + 1
    GostergeYaz ws, satir, "Hızlı kazanım", hizli, "0": satir = satir + 1
    If sonuclanan > 0 Then
        GostergeYaz ws, satir, "Kabul oranı", kabul / sonuclanan, "0%"
    Else
        GostergeYaz ws, satir, "Kabul oranı", 0, "0%"
    End If
    satir = satir + 2

    BaslikYaz ws, satir, "Gerçekleşen kazanım"
    satir = satir + 1
    GostergeYaz ws, satir, "Yıllık kazanılan saat", saat, "#,##0.0": satir = satir + 1
    GostergeYaz ws, satir, "Yıllık TL tasarrufu", tl, "#,##0 ₺": satir = satir + 2

    With ws.Range(ws.Cells(satir, 2), ws.Cells(satir + 1, 3))
        .Merge
        .Value = "Kazanım rakamlarına yalnızca uygulamaya geçmiş öneriler " & _
                 "(Pilot Uygulamada, Ölçümleniyor, Standartlaştırıldı) dahil " & _
                 "edilmiştir. Henüz uygulanmamış öneriler tasarruf olarak sayılmaz."
        .Font.Name = modTasarim.FONT_AILE
        .Font.Size = 9
        .Font.Color = modTasarim.HexRGB(modTasarim.CLR_METIN_GRI)
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = modTasarim.HexRGB(modTasarim.CLR_ACIK_MAVI)
    End With
    ws.Rows(satir & ":" & satir + 1).RowHeight = 20

    ' --- Durum dagilimi ---------------------------------------------------
    satir = satir + 3
    BaslikYaz ws, satir, "Durum dağılımı"
    satir = satir + 1

    Dim liste As Variant, d As Long, sayac As Long
    liste = modModel.Durumlar()
    For d = LBound(liste) To UBound(liste)
        sayac = 0
        For k = 1 To adet
            If StrComp(CStr(veri(secili(k), modKonsolide.V_DURUM)), _
                       CStr(liste(d)), vbTextCompare) = 0 Then sayac = sayac + 1
        Next k
        ws.Cells(satir, 2).Value = CStr(liste(d))
        ws.Cells(satir, 3).Value = sayac
        ws.Cells(satir, 3).NumberFormat = "0"
        modTasarim.DurumRozetiUygula ws.Cells(satir, 2), CStr(liste(d))
        ws.Cells(satir, 2).HorizontalAlignment = xlLeft
        ws.Cells(satir, 2).IndentLevel = 1
        ws.Cells(satir, 3).Font.Name = modTasarim.FONT_AILE
        satir = satir + 1
    Next d

    ws.Range("B:C").VerticalAlignment = xlCenter
End Sub

Private Sub BaslikYaz(ByVal ws As Object, ByVal satir As Long, ByVal metin As String)
    With ws.Cells(satir, 2)
        .Value = metin
        .Font.Name = modTasarim.FONT_BASLIK_AILE
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = modTasarim.HexRGB(modTasarim.CLR_ANA_LACIVERT)
    End With
    ws.Rows(satir).RowHeight = 22
End Sub

Private Sub GostergeYaz(ByVal ws As Object, ByVal satir As Long, ByVal etiket As String, _
                        ByVal deger As Variant, ByVal bicim As String)
    With ws.Cells(satir, 2)
        .Value = etiket
        .Font.Name = modTasarim.FONT_AILE
        .Font.Size = 10.5
        .Font.Color = modTasarim.HexRGB(modTasarim.CLR_METIN_KOYU)
        .IndentLevel = 1
    End With
    With ws.Cells(satir, 3)
        .Value = deger
        .NumberFormat = bicim
        .Font.Name = modTasarim.FONT_BASLIK_AILE
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = modTasarim.HexRGB(modTasarim.CLR_ANA_LACIVERT)
        .HorizontalAlignment = xlRight
    End With
    ws.Cells(satir, 2).Resize(1, 2).Borders(xlEdgeBottom).Color = _
        modTasarim.HexRGB(modTasarim.CLR_CIZGI_GRI)
    ws.Rows(satir).RowHeight = 20
End Sub


' --------------------------------------------------------------------------
'  Oneriler sayfasi -- deger kopyasi (formul yok, bagimlilik yok)
' --------------------------------------------------------------------------
Private Sub TabloYaz(ByVal ws As Object, ByVal veri As Variant, _
                     ByRef secili() As Long, ByVal adet As Long)
    Dim basliklar As Variant
    Dim tablo() As Variant, k As Long, i As Long, s As Long

    basliklar = Array("Öneri No", "Tarih", "Gönderen", "Öneri Başlığı", _
                      "Durum", "PDCA", "Etki", "Efor", "Öncelik", _
                      "Yıllık Saat", "Yıllık TL", "Son Değerlendiren", "Karar Notu")

    For s = LBound(basliklar) To UBound(basliklar)
        ws.Cells(1, s + 1).Value = basliklar(s)
    Next s

    With ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(basliklar) + 1))
        .Interior.Color = modTasarim.HexRGB(modTasarim.CLR_ANA_LACIVERT)
        .Font.Color = modTasarim.HexRGB(modTasarim.CLR_BEYAZ)
        .Font.Name = modTasarim.FONT_AILE
        .Font.Bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Rows(1).RowHeight = 28

    ReDim tablo(1 To adet, 1 To UBound(basliklar) + 1)
    For k = 1 To adet
        i = secili(k)
        tablo(k, 1) = veri(i, modKonsolide.V_ONERI_NO)
        tablo(k, 2) = veri(i, modKonsolide.V_TARIH)
        tablo(k, 3) = veri(i, modKonsolide.V_AD_SOYAD)
        tablo(k, 4) = veri(i, modKonsolide.V_BASLIK)
        tablo(k, 5) = veri(i, modKonsolide.V_DURUM)
        tablo(k, 6) = modModel.PDCA(CStr(veri(i, modKonsolide.V_DURUM)))
        tablo(k, 7) = modKonsolide.BosDegilse(veri(i, modKonsolide.V_ETKI))
        tablo(k, 8) = modKonsolide.BosDegilse(veri(i, modKonsolide.V_EFOR))
        tablo(k, 9) = veri(i, modKonsolide.V_ONCELIK)
        tablo(k, 10) = modKonsolide.BosDegilse(veri(i, modKonsolide.V_SAAT))
        tablo(k, 11) = modKonsolide.BosDegilse(veri(i, modKonsolide.V_TL))
        tablo(k, 12) = veri(i, modKonsolide.V_DEGERLENDIREN)
        tablo(k, 13) = veri(i, modKonsolide.V_KARAR_NOTU)
    Next k

    ws.Cells(2, 1).Resize(adet, UBound(basliklar) + 1).Value = tablo

    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 22
    ws.Columns("D").ColumnWidth = 46
    ws.Columns("E").ColumnWidth = 19
    ws.Columns("F").ColumnWidth = 10
    ws.Columns("G:H").ColumnWidth = 7
    ws.Columns("I").ColumnWidth = 20
    ws.Columns("J:K").ColumnWidth = 13
    ws.Columns("L").ColumnWidth = 18
    ws.Columns("M").ColumnWidth = 46

    With ws.Range(ws.Cells(2, 1), ws.Cells(adet + 1, UBound(basliklar) + 1))
        .Font.Name = modTasarim.FONT_AILE
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .Borders(xlInsideHorizontal).Color = modTasarim.HexRGB(modTasarim.CLR_CIZGI_GRI)
        .Borders(xlInsideVertical).Color = modTasarim.HexRGB(modTasarim.CLR_CIZGI_GRI)
    End With
    ws.Range(ws.Cells(2, 10), ws.Cells(adet + 1, 10)).NumberFormat = "#,##0.0"
    ws.Range(ws.Cells(2, 11), ws.Cells(adet + 1, 11)).NumberFormat = "#,##0 ₺"

    ' Durum ve oncelik sutunlarini rozetle
    For k = 1 To adet
        modTasarim.DurumRozetiUygula ws.Cells(k + 1, 5), CStr(tablo(k, 5))
        If Len(CStr(tablo(k, 9))) > 0 Then
            modTasarim.OncelikRozetiUygula ws.Cells(k + 1, 9), CStr(tablo(k, 9))
        End If
    Next k

    ws.Range(ws.Cells(1, 1), ws.Cells(adet + 1, UBound(basliklar) + 1)).AutoFilter
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub


' ###########################################################################
'  YARDIMCILAR
' ###########################################################################

Public Function RaporSayfasi() As Object
    Set RaporSayfasi = ThisWorkbook.Worksheets(modUI.SAYFA_RAPOR)
End Function

Private Function Sayi(ByVal v As Variant) As Double
    If IsNumeric(v) Then Sayi = CDbl(v) Else Sayi = 0
End Function


' ###########################################################################
'  TESTLER ICIN -- ekran olmadan rapor uretimi
' ###########################################################################
Public Function TestRaporu(ByVal donem As String, ByVal parola As String) As String
    Dim veri As Variant, secili() As Long, adet As Long

    veri = modKonsolide.VeriDizisi()
    If IsEmpty(veri) Then
        TestRaporu = ""
        Exit Function
    End If

    secili = DonemSuz(veri, donem, adet)
    If adet = 0 Then
        TestRaporu = ""
        Exit Function
    End If

    TestRaporu = KitabiUret(veri, secili, adet, donem, parola)
End Function
