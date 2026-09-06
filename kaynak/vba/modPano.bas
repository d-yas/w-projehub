Attribute VB_Name = "modPano"
Option Explicit

' ============================================================================
'  modPano -- gostergeler ve grafik verileri
'
'  Gostergeler hucre formulleriyle degil VBA ile hesaplanir. Gerekce:
'  bekleyen ve uygulanmis tanimlari modModel'de tek yerde durur. Ayni kurali
'  bir de formul olarak yazmak, ikisinin zamanla birbirinden ayrilmasi
'  demektir.
' ============================================================================

' PanoVeri sayfasindaki grafik kaynak bloklari
Private Const PV_DURUM_SATIR As Long = 2       ' A2:B9   durum, adet
Private Const PV_AY_SATIR As Long = 2          ' D2:E13  ay, adet
Private Const PV_AY_SAYISI As Long = 12


' ###########################################################################
'  DUGME EYLEMLERI
' ###########################################################################

Public Sub PanoyuYenile()
    On Error GoTo Hata
    Hesapla
    Exit Sub
Hata:
    modUI.Hata "Pano hesaplanamadı." & vbCrLf & vbCrLf & Err.Description, "Pano"
End Sub

Public Sub PanoyaGit()
    modUI.SayfaGoster modUI.SAYFA_PANO
    PanoyuYenile
End Sub


' ###########################################################################
'  HESAPLAMA
' ###########################################################################

Public Sub Hesapla()
    Dim ws As Object, veri As Variant
    Dim n As Long, i As Long, durum As String

    Dim toplam As Long, bekleyen As Long, uygulanan As Long, buAy As Long
    Dim durumSayac As Object, aySayac As Object

    Set ws = PanoSayfasi()
    veri = modKonsolide.VeriDizisi()

    Set durumSayac = CreateObject("Scripting.Dictionary")
    Set aySayac = CreateObject("Scripting.Dictionary")
    durumSayac.CompareMode = 1
    aySayac.CompareMode = 1

    If Not IsEmpty(veri) Then
        n = UBound(veri, 1)
        For i = 1 To n
            durum = CStr(veri(i, modKonsolide.V_DURUM))
            toplam = toplam + 1

            If modModel.BekliyorMu(durum) Then bekleyen = bekleyen + 1
            If modModel.UygulanmisMi(durum) Then uygulanan = uygulanan + 1

            durumSayac(durum) = SayiAl(durumSayac(durum)) + 1

            Dim ayAnahtar As String
            ayAnahtar = Left$(CStr(veri(i, modKonsolide.V_TARIH)), 7)   ' yyyy-mm
            If Len(ayAnahtar) = 7 Then
                aySayac(ayAnahtar) = SayiAl(aySayac(ayAnahtar)) + 1
                If ayAnahtar = Format$(Now, "yyyy-mm") Then buAy = buAy + 1
            End If
        Next i
    End If

    modUI.KorumaKapa ws
    ws.Range("pano_toplam").Value = toplam
    ws.Range("pano_bekleyen").Value = bekleyen
    ws.Range("pano_uygulanan").Value = uygulanan
    ws.Range("pano_bu_ay").Value = buAy

    ws.Range("pano_guncelleme").Value = _
        "Son güncelleme: " & Format$(Now, "dd.mm.yyyy hh:nn")
    modUI.KorumaAc ws

    KartlariTazele toplam, bekleyen, uygulanan, buAy
    GrafikVerisiYaz durumSayac, aySayac
End Sub


' ###########################################################################
'  KART SEKILLERI
'
'  Panodaki kartlarin govdesi yuvarlak koseli birer SEKILDIR (uretim
'  asamasinda com_kurulum.kpi_karti_ciz ekler). Excel'de sekiller her zaman
'  hucrelerin ustunde cizildigi icin rakam da sekle tasinmak zorunda kaldi.
'
'  Gostergenin TEK DOGRULUK KAYNAGI yine hucredir: yukaridaki Hesapla oraya
'  yazar, testler oradan okur. Burasi yalnizca gorunumu esitler. Sekil bir
'  sebeple yoksa (eski dosya, uretim sirasinda atlanmis sekil katmani)
'  gosterge yine dogru hesaplanmis olur.
' ###########################################################################

Private Sub KartlariTazele(ByVal toplam As Long, ByVal bekleyen As Long, _
                           ByVal uygulanan As Long, ByVal buAy As Long)
    KartYaz "pano_toplam", toplam
    KartYaz "pano_bekleyen", bekleyen
    KartYaz "pano_uygulanan", uygulanan
    KartYaz "pano_bu_ay", buAy
End Sub

Private Sub KartYaz(ByVal ad As String, ByVal deger As Long)
    On Error Resume Next
    PanoSayfasi().Shapes("kpi_deger_" & ad).TextFrame2.TextRange.Text = _
        Format$(deger, "#,##0")
    On Error GoTo 0
End Sub


' ###########################################################################
'  GRAFIK KAYNAK VERISI
' ###########################################################################

Private Sub GrafikVerisiYaz(ByVal durumSayac As Object, ByVal aySayac As Object)
    Dim ws As Object
    Dim liste As Variant, i As Long
    Dim ayAnahtar As String, ilkAy As Date

    Set ws = PanoVeriSayfasi()
    modUI.KorumaKapa ws

    ' --- Durum dagilimi ---------------------------------------------------
    liste = modModel.Durumlar()
    For i = LBound(liste) To UBound(liste)
        ws.Cells(PV_DURUM_SATIR + i, 1).Value = liste(i)
        ws.Cells(PV_DURUM_SATIR + i, 2).Value = SayiAl(durumSayac(CStr(liste(i))))
    Next i

    ' --- Son 12 ayin trendi ----------------------------------------------
    ' Ay etiketleri her yenilemede bugune gore yeniden uretilir; grafik
    ' aralik sabit kaldigi icin ayrica guncellenmesi gerekmez.
    ilkAy = DateSerial(Year(Now), Month(Now) - (PV_AY_SAYISI - 1), 1)
    For i = 0 To PV_AY_SAYISI - 1
        ayAnahtar = Format$(DateAdd("m", i, ilkAy), "yyyy-mm")
        ws.Cells(PV_AY_SATIR + i, 4).Value = Format$(DateAdd("m", i, ilkAy), "mmm yy")
        ws.Cells(PV_AY_SATIR + i, 5).Value = SayiAl(aySayac(ayAnahtar))
    Next i

    modUI.KorumaAc ws
End Sub


' ###########################################################################
'  YARDIMCILAR
' ###########################################################################

Public Function PanoSayfasi() As Object
    Set PanoSayfasi = ThisWorkbook.Worksheets(modUI.SAYFA_PANO)
End Function

Public Function PanoVeriSayfasi() As Object
    Set PanoVeriSayfasi = ThisWorkbook.Worksheets(modUI.SAYFA_PANOVERI)
End Function

Private Function SayiAl(ByVal v As Variant) As Double
    If IsNumeric(v) Then SayiAl = CDbl(v) Else SayiAl = 0
End Function

' ###########################################################################
'  TESTLER ICIN -- gosterge degerlerini tek dizgede dondurur
' ###########################################################################
Public Function TestGostergeleri() As String
    Dim ws As Object
    Hesapla
    Set ws = PanoSayfasi()
    TestGostergeleri = _
        "toplam=" & ws.Range("pano_toplam").Value & ";" & _
        "bekleyen=" & ws.Range("pano_bekleyen").Value & ";" & _
        "uygulanan=" & ws.Range("pano_uygulanan").Value & ";" & _
        "bu_ay=" & ws.Range("pano_bu_ay").Value
End Function

' Bir kart sekli uzerinde gorunen metni dondurur (bos dizge = sekil yok).
Public Function TestKartMetni(ByVal ad As String) As String
    On Error Resume Next
    TestKartMetni = PanoSayfasi().Shapes("kpi_deger_" & ad) _
                        .TextFrame2.TextRange.Text
    On Error GoTo 0
End Function
