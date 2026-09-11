Attribute VB_Name = "modTakip"
Option Explicit

' NOT: Modul duzeyi bildirimler VBA'da YALNIZCA burada, ilk yordamdan once
' bulunabilir (test_uretim.py bunu denetler).

' --- Durum gecmisi tablosu (uret_takip.py ile ayni olmali) ----------------
Private Const GECMIS_ILK_SATIR As Long = 31
Private Const GECMIS_AZAMI As Long = 10
Private Const GECMIS_ILK_SUTUN As Long = 2     ' B: tarih, C: durum, D:F aciklama
' F dahil: aciklama sutunu D:F birlesiktir ve ClearContents birlesimin yalnizca
' bir parcasi uzerinde calistirilamaz.
Private Const GECMIS_SON_SUTUN As Long = 6     ' F

' ============================================================================
'  modTakip -- personelin oneri takip ekrani (ProjeTakip.xlsm)
'
'  Personel oneri numarasini ve sicil numarasini yazar, "Sorgula"ya basar;
'  ekrana o onerinin guncel durumu ve durum gecmisi gelir.
'
'  GIZLILIK -- bu ekranin uc kurali
'
'  1) DEPO BU KITABA KOPYALANMAZ. Takip kitabi parolasizdir ve personelin
'     elindedir. Depo gizli bir Excel orneginde okunur ve buraya yalnizca
'     sorgulanan onerinin satirlari gelir (modDepo.OneriSatirlariniOku).
'
'  2) NUMARA TEK BASINA YETMEZ. Numaralar sirali oldugu icin tahmin
'     edilebilir; sonuc yalnizca gonderimdeki sicil numarasi da tutarsa
'     gosterilir. "Boyle bir numara yok" ile "sicil tutmadi" AYNI mesajla
'     soylenir -- aksi halde ekran hangi numaralarin var oldugunu ele verirdi.
'
'  3) EKIBIN NOTLARI GOSTERILMEZ. Degerlendirme ekibi karar notunu ic kayit
'     olarak yazar; oneri sahibi durumu, tarihleri ve durumun anlamini gorur.
'     Degerlendirenin kim oldugu da gosterilmez. test_uretim.py bu modulun o
'     sutunlara hic dokunmadigini denetler.
' ============================================================================


' ###########################################################################
'  DUGME EYLEMLERI
' ###########################################################################

' "Sorgula" dugmesi
Public Sub Sorgula()
    Dim ws As Object
    Dim no As String, sicil As String, hataMetni As String
    Dim oneri As Variant, olaylar As Variant

    Set ws = TakipSayfasi()
    no = NoDuzelt(ws.Range("tk_no").Value)
    sicil = Trim$(CStr(ws.Range("tk_sicil").Value & ""))

    ' Onceki sorgunun sonucu HER DURUMDA silinir: yeni sorgu sonuc vermezse
    ' ekranda baska bir onerinin durumu kalmamalidir.
    SonucuTemizle ws

    If Len(no) = 0 Or Len(sicil) = 0 Then
        Uyar ws, "Öneri numarası ve sicil numarası birlikte girilmelidir.", _
             "Öneri numarası ve sicil numarası birlikte girilmelidir."
        Exit Sub
    End If

    modUI.BantYaz ws, "tk_bant", "Sorgulanıyor…", _
                  modTasarim.CLR_ACIK_MAVI, modTasarim.CLR_ANA_LACIVERT
    DoEvents

    On Error GoTo Hata
    modUI.HizliModAc

    modDepo.OneriSatirlariniOku no, oneri, olaylar

    ' "Numara yok" ile "sicil tutmadi" ayni yoldan gecer; bkz. modul basi.
    If Not SicilTutuyorMu(oneri, sicil) Then
        modUI.HizliModKapa
        Uyar ws, "Bu numara ve sicil numarasıyla eşleşen bir öneri bulunamadı.", _
             "Bu öneri numarası ve sicil numarasıyla eşleşen bir öneri " & _
             "bulunamadı." & vbCrLf & vbCrLf & _
             "Öneri numarasını gönderimde verildiği gibi (örnek: " & _
             modAyar.ONEK_ONERI_NO & "-" & Year(Now) & "-0001), sicil " & _
             "numarasını öneriyi gönderirken yazdığınız gibi girin."
        Exit Sub
    End If

    SonucuYaz ws, oneri, olaylar
    modUI.HizliModKapa

    modUI.BantYaz ws, "tk_bant", _
        modTasarim.IsaretOnay() & "  Öneriniz bulundu  ·  sorgu zamanı " & Format$(Now, "dd.mm.yyyy hh:nn"), _
        modTasarim.CLR_ONAY_ZEMIN, modTasarim.CLR_ONAY_YAZI
    Exit Sub

Hata:
    ' Aciklama ILK is olarak alinir: sonraki her cagri Err'i temizler.
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    modUI.HizliModKapa
    HataBandi ws
    modUI.Hata "Önerinizin durumu okunamadı." & vbCrLf & vbCrLf & hataMetni & _
               vbCrLf & vbCrLf & _
               "Birkaç saniye sonra yeniden deneyin. Sorun sürerse " & _
               "Değerlendirme ekibine başvurun.", "Öneri Takibi"
End Sub

' "Temizle" dugmesi
Public Sub Temizle()
    EkraniSifirla
End Sub

' Girisleri, sonucu ve bandi temizler. Kitap her acildiginda da cagrilir:
' onceki sorgunun sicil numarasi ve sonucu ekranda kalmamalidir.
Public Sub EkraniSifirla()
    Dim ws As Object, hataMetni As String

    Set ws = TakipSayfasi()

    On Error GoTo Hata
    SonucuTemizle ws
    modUI.KorumaKapa ws
    modUI.Alan(ws, "tk_no").ClearContents
    modUI.Alan(ws, "tk_sicil").ClearContents
    modUI.KorumaAc ws
    modUI.BantTemizle ws, "tk_bant"

    ' Imleci ilk alana goturmek gorsel bir kolayliktir; penceresi olmayan bir
    ' oturumda basarisiz olabilir ve bu onemli degildir.
    On Error Resume Next
    ws.Activate
    ws.Range("tk_no").Select
    On Error GoTo 0
    Exit Sub

Hata:
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    HataBandi ws
    modUI.Hata "Ekran temizlenemedi." & vbCrLf & vbCrLf & hataMetni, "Öneri Takibi"
End Sub


' ###########################################################################
'  EKRAN
' ###########################################################################

Private Sub SonucuYaz(ByVal ws As Object, ByVal oneri As Variant, ByVal olaylar As Variant)
    Dim tarihler() As String, durumlar() As String
    Dim durum As String, yeni As String
    Dim n As Long, r As Long, i As Long, satir As Long, gosterilecek As Long

    ' --- Durum gecmisi: gonderim + DURUMU DEGISTIREN olaylar ------------
    ' Yalnizca not ekleyen olaylar (durum ayni kalir) gosterilmez: notlar bu
    ' ekranda yoktur ve ayni durumun art arda iki kez gorunmesi "bir sey
    ' degisti" sanilirdi. Kural listeyle aynidir (modModel.OlaySonrasiDurum).
    n = 1
    ReDim tarihler(1 To n)
    ReDim durumlar(1 To n)
    tarihler(1) = Trim$(CStr(oneri(1, modDepo.O_TARIH) & ""))
    durumlar(1) = modModel.DURUM_YENI
    durum = modModel.DURUM_YENI

    If Not IsEmpty(olaylar) Then
        For r = LBound(olaylar, 1) To UBound(olaylar, 1)
            yeni = modModel.OlaySonrasiDurum(durum, olaylar(r, modDepo.E_YENI_DURUM))
            If StrComp(yeni, durum, vbTextCompare) <> 0 Then
                n = n + 1
                ReDim Preserve tarihler(1 To n)
                ReDim Preserve durumlar(1 To n)
                tarihler(n) = Trim$(CStr(olaylar(r, modDepo.E_OLAY_TARIHI) & ""))
                durumlar(n) = yeni
                durum = yeni
            End If
        Next r
    End If

    modUI.KorumaKapa ws
    ws.Range("tk_sonuc_no").Value = CStr(oneri(1, modDepo.O_ONERI_NO) & "")
    ws.Range("tk_tarih").Value = TarihGoster(tarihler(1), False)
    ws.Range("tk_baslik").Value = CStr(oneri(1, modDepo.O_ONERI_BASLIGI) & "")
    ws.Range("tk_durum").Value = durum
    ws.Range("tk_guncelleme").Value = TarihGoster(tarihler(n), False)

    ' En yeni degisiklik ustte. Tabloya sigmayan eski degisikliklerin yerine
    ' son satir toplami soyler.
    gosterilecek = n
    If n > GECMIS_AZAMI Then gosterilecek = GECMIS_AZAMI - 1

    satir = GECMIS_ILK_SATIR
    For i = n To n - gosterilecek + 1 Step -1
        ws.Cells(satir, GECMIS_ILK_SUTUN).Value = TarihGoster(tarihler(i), True)
        ws.Cells(satir, GECMIS_ILK_SUTUN + 1).Value = durumlar(i)
        ws.Cells(satir, GECMIS_ILK_SUTUN + 2).Value = modModel.DurumAciklamasi(durumlar(i))
        satir = satir + 1
    Next i

    If n > GECMIS_AZAMI Then
        ws.Cells(satir, GECMIS_ILK_SUTUN + 2).Value = _
            "… toplam " & n & " değişiklik; daha eskileri gösterilmiyor."
    End If
    modUI.KorumaAc ws
End Sub

Private Sub SonucuTemizle(ByVal ws As Object)
    Dim ad As Variant

    modUI.KorumaKapa ws
    For Each ad In Array("tk_sonuc_no", "tk_tarih", "tk_baslik", "tk_durum", "tk_guncelleme")
        modUI.Alan(ws, CStr(ad)).ClearContents
    Next ad
    ws.Range(ws.Cells(GECMIS_ILK_SATIR, GECMIS_ILK_SUTUN), _
             ws.Cells(GECMIS_ILK_SATIR + GECMIS_AZAMI - 1, GECMIS_SON_SUTUN)).ClearContents
    modUI.KorumaAc ws
End Sub

' Hem kullaniciya mesaj verir hem de bandi uyari rengine boyar. Bant tek
' satirliktir; mesaj kutusu ayrintiyi tasiyabilir.
Private Sub Uyar(ByVal ws As Object, ByVal bantMetni As String, ByVal mesaj As String)
    modUI.BantYaz ws, "tk_bant", modTasarim.IsaretUyari() & "  " & bantMetni, _
                  modTasarim.CLR_UYARI_ZEMIN, modTasarim.CLR_UYARI_YAZI
    modUI.Hata mesaj, "Öneri Takibi"
End Sub

' Hata isleyicilerinden cagrilir. VBA'da isleyicinin icinde olusan hata
' YAKALANAMAZ; bu yuzden kendi icinde tamamen hataya dayaniklidir.
Private Sub HataBandi(ByVal ws As Object)
    On Error Resume Next
    SonucuTemizle ws
    modUI.BantYaz ws, "tk_bant", modTasarim.IsaretUyari() & "  Sorgu yapılamadı.", _
                  modTasarim.CLR_UYARI_ZEMIN, modTasarim.CLR_UYARI_YAZI
    modUI.KorumaAc ws
    Err.Clear
    On Error GoTo 0
End Sub


' ###########################################################################
'  YARDIMCILAR
' ###########################################################################

Private Function TakipSayfasi() As Object
    Set TakipSayfasi = ThisWorkbook.Worksheets(modUI.SAYFA_TAKIP)
End Function

' Kullanicilar numarayi kucuk harfle ya da bosluklu yazar: " prj-2026-0001".
Private Function NoDuzelt(ByVal deger As Variant) As String
    NoDuzelt = Replace(UCase$(Trim$(CStr(deger & ""))), " ", "")
End Function

Private Function SicilTutuyorMu(ByVal oneri As Variant, ByVal sicil As String) As Boolean
    If IsEmpty(oneri) Then Exit Function
    SicilTutuyorMu = (StrComp(Trim$(CStr(oneri(1, modDepo.O_SICIL_NO) & "")), _
                              sicil, vbTextCompare) = 0)
End Function

' "2026-09-04T10:11:51" -> "04.09.2026" (saatli: "04.09.2026  10:11")
' Kayitta tarih ISO bicimindedir (siralanabilir olsun diye); ekranda okunur
' bicimde gosterilir.
Private Function TarihGoster(ByVal isoTarih As String, ByVal saatli As Boolean) As String
    If Len(isoTarih) < 10 Then
        TarihGoster = isoTarih
        Exit Function
    End If

    TarihGoster = Mid$(isoTarih, 9, 2) & "." & Mid$(isoTarih, 6, 2) & "." & _
                  Left$(isoTarih, 4)
    If saatli And Len(isoTarih) >= 16 Then
        TarihGoster = TarihGoster & "  " & Mid$(isoTarih, 12, 5)
    End If
End Function
