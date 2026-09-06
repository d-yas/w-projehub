Attribute VB_Name = "modPano"
Option Explicit

' ============================================================================
'  modPano -- gostergeler ve grafik verileri
'
'  Gostergeler hucre formulleriyle degil VBA ile hesaplanir. Gerekce:
'  tasarruf kurali, kabul tanimi ve oncelik siniflari modModel'de tek yerde
'  durur. Ayni kurali bir de formul olarak yazmak, ikisinin zamanla
'  birbirinden ayrilmasi demektir.
'
'  EN ONEMLI KURAL -- TASARRUF:
'  Yillik saat ve TL toplamlarina yalnizca fayda gerceklesmis sayilan
'  durumlardaki oneriler girer (Pilot Uygulamada, Ölçümleniyor,
'  Standartlaştırıldı). Heniz uygulanmamis bir oneri tasarruf olarak
'  sayilsaydi pano gercekte olmayan bir kazanci yonetime raporlardi.
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

    Dim toplam As Long, bekleyen As Long, uygulanan As Long, kabul As Long
    Dim sonuclanan As Long, hizli As Long, buAy As Long
    Dim saat As Double, tl As Double
    Dim yanitGun As Double, yanitAdet As Long

    Dim durumSayac As Object, aySayac As Object
    Dim mtx(1 To 4) As Long
    Dim sinif As String

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
            If modModel.KabulEdildiMi(durum) Then kabul = kabul + 1
            If modModel.SonuclandiMi(durum) Then sonuclanan = sonuclanan + 1
            If modModel.TasarrufSayilirMi(durum) Then
                uygulanan = uygulanan + 1
                saat = saat + SayiAl(veri(i, modKonsolide.V_SAAT))
                tl = tl + SayiAl(veri(i, modKonsolide.V_TL))
            End If

            sinif = CStr(veri(i, modKonsolide.V_ONCELIK))
            If StrComp(sinif, modModel.ONCELIK_HIZLI, vbTextCompare) = 0 Then
                hizli = hizli + 1
                mtx(1) = mtx(1) + 1
            ElseIf StrComp(sinif, modModel.ONCELIK_BUYUK, vbTextCompare) = 0 Then
                mtx(2) = mtx(2) + 1
            ElseIf StrComp(sinif, modModel.ONCELIK_DOLDURMA, vbTextCompare) = 0 Then
                mtx(3) = mtx(3) + 1
            ElseIf StrComp(sinif, modModel.ONCELIK_DISI, vbTextCompare) = 0 Then
                mtx(4) = mtx(4) + 1
            End If

            durumSayac(durum) = SayiAl(durumSayac(durum)) + 1

            Dim ayAnahtar As String
            ayAnahtar = Left$(CStr(veri(i, modKonsolide.V_TARIH)), 7)   ' yyyy-mm
            If Len(ayAnahtar) = 7 Then
                aySayac(ayAnahtar) = SayiAl(aySayac(ayAnahtar)) + 1
                If ayAnahtar = Format$(Now, "yyyy-mm") Then buAy = buAy + 1
            End If

            ' Ilk yanit suresi: gonderim ile ILK degerlendirme olayi arasi.
            Dim g As Date, o As Date
            If TarihCoz(CStr(veri(i, modKonsolide.V_TARIH)), g) And _
               TarihCoz(CStr(veri(i, modKonsolide.V_ILK_OLAY)), o) Then
                If o >= g Then
                    yanitGun = yanitGun + (o - g)
                    yanitAdet = yanitAdet + 1
                End If
            End If
        Next i
    End If

    modUI.KorumaKapa ws
    ws.Range("pano_toplam").Value = toplam
    ws.Range("pano_bekleyen").Value = bekleyen
    ws.Range("pano_uygulanan").Value = uygulanan
    ws.Range("pano_saat").Value = saat
    ws.Range("pano_tl").Value = tl
    ws.Range("pano_hizli").Value = hizli

    If sonuclanan > 0 Then
        ws.Range("pano_kabul_orani").Value = kabul / sonuclanan
    Else
        ws.Range("pano_kabul_orani").Value = 0
    End If

    If yanitAdet > 0 Then
        ws.Range("pano_yanit_suresi").Value = yanitGun / yanitAdet
    Else
        ws.Range("pano_yanit_suresi").Value = 0
    End If

    ws.Range("pano_bu_ay").Value = buAy

    ws.Range("pano_mtx_hizli").Value = mtx(1)
    ws.Range("pano_mtx_buyuk").Value = mtx(2)
    ws.Range("pano_mtx_doldurma").Value = mtx(3)
    ws.Range("pano_mtx_disi").Value = mtx(4)

    ws.Range("pano_guncelleme").Value = _
        "Son güncelleme: " & Format$(Now, "dd.mm.yyyy hh:nn") & _
        "   ·   Tasarruf yalnızca uygulamaya geçmiş önerilerden sayılır."
    modUI.KorumaAc ws

    GrafikVerisiYaz durumSayac, aySayac
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

' "2026-09-04T10:11:51" -> Date. Bicim bozuksa False doner.
' CDate kullanilmaz: makinenin tarih bicimi ayarina bagimli olurdu.
Private Function TarihCoz(ByVal isoTarih As String, ByRef sonuc As Date) As Boolean
    Dim y As Long, a As Long, g As Long
    If Len(isoTarih) < 10 Then Exit Function
    If Not IsNumeric(Left$(isoTarih, 4)) Then Exit Function
    If Not IsNumeric(Mid$(isoTarih, 6, 2)) Then Exit Function
    If Not IsNumeric(Mid$(isoTarih, 9, 2)) Then Exit Function

    y = CLng(Left$(isoTarih, 4))
    a = CLng(Mid$(isoTarih, 6, 2))
    g = CLng(Mid$(isoTarih, 9, 2))
    If a < 1 Or a > 12 Or g < 1 Or g > 31 Then Exit Function

    On Error GoTo Hata
    sonuc = DateSerial(y, a, g)
    TarihCoz = True
    Exit Function
Hata:
    TarihCoz = False
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
        "saat=" & Replace(CStr(ws.Range("pano_saat").Value), ",", ".") & ";" & _
        "tl=" & Replace(CStr(ws.Range("pano_tl").Value), ",", ".") & ";" & _
        "hizli=" & ws.Range("pano_hizli").Value & ";" & _
        "mtx_hizli=" & ws.Range("pano_mtx_hizli").Value & ";" & _
        "mtx_buyuk=" & ws.Range("pano_mtx_buyuk").Value & ";" & _
        "mtx_doldurma=" & ws.Range("pano_mtx_doldurma").Value & ";" & _
        "mtx_disi=" & ws.Range("pano_mtx_disi").Value
End Function
