Attribute VB_Name = "modDegerlendirme"
Option Explicit

' ============================================================================
'  modDegerlendirme -- degerlendirme ekrani ve EKLE-ONLY kayit
'
'  Hicbir kayit guncellenmez, hicbir dosya silinmez. Her degerlendirme
'  "degerlendirme\<yil>\" altina YENI bir olay dosyasi birakir. Bir onerinin
'  gecmisi bu dosyalarin toplamidir; "kim, ne zaman, neyi degistirdi" izi
'  yapisal olarak korunur -- silinebilecek bir gecmis alani yoktur.
'
'  Dosya adi: <oneri_no>_<YYYYMMDDHHMMSSss>_<rastgele>.txt
'  Ada gore siralama = zaman sirasi. Ayni saniyede iki ekip uyesi kaydetse
'  bile rastgele ek sayesinde birbirlerinin dosyasini ezmezler.
' ============================================================================

Private Const GECMIS_ILK_SATIR As Long = 34
Private Const GECMIS_AZAMI As Long = 14
Private Const GECMIS_ILK_SUTUN As Long = 2     ' B
' F dahil: not sutunu E:F birlesiktir ve ClearContents birlesimin yalnizca
' bir parcasi uzerinde calistirilamaz.
Private Const GECMIS_SON_SUTUN As Long = 6     ' F

Private Function OlayAlanlari() As Variant
    OlayAlanlari = Array("sema", "oneri_no", "olay_tarihi", _
                         "yeni_durum", "karar_notu", _
                         "degerlendiren_kullanici", "degerlendiren_bilgisayar")
End Function


' ###########################################################################
'  DUGME EYLEMLERI
' ###########################################################################

' Listedeki secili satiri degerlendirme ekraninda acar.
Public Sub SeciliyiDegerlendir()
    Dim ws As Object, no As String

    Set ws = modKonsolide.ListeSayfasi()
    If StrComp(ActiveSheet.Name, ws.Name, vbTextCompare) <> 0 Then
        modUI.SayfaGoster modUI.SAYFA_LISTE
        modUI.Bilgi "Önce listeden bir öneri satırı seçin, sonra " & _
                    "“Seçiliyi Değerlendir” düğmesine basın.", "Değerlendirme"
        Exit Sub
    End If

    no = modKonsolide.SeciliOneriNo()
    If Len(no) = 0 Then
        modUI.Bilgi "Listede bir öneri satırı seçili değil." & vbCrLf & vbCrLf & _
                    "Değerlendirmek istediğiniz önerinin satırına tıklayın, " & _
                    "sonra bu düğmeye basın.", "Değerlendirme"
        Exit Sub
    End If

    OneriyiAc no
End Sub

Public Sub ListeyeDon()
    modUI.SayfaGoster modUI.SAYFA_LISTE
End Sub

' Degerlendirmeyi yeni bir olay dosyasi olarak kaydeder.
Public Sub DegerlendirmeKaydet()
    Dim ws As Object, sozluk As Object
    Dim no As String, hataMetni As String, hedefDosya As String

    Set ws = DegerlendirmeSayfasi()
    no = Trim$(CStr(ws.Range("dg_oneri_no").Value & ""))

    If Len(no) = 0 Then
        modUI.Hata "Ekranda açık bir öneri yok." & vbCrLf & vbCrLf & _
                   "Listeden bir öneri seçip “Seçiliyi Değerlendir” " & _
                   "düğmesine basın.", "Değerlendirme"
        Exit Sub
    End If

    Set sozluk = EkrandanOku(ws, no)
    hataMetni = Dogrula(sozluk)
    If Len(hataMetni) > 0 Then
        modUI.Hata hataMetni, "Eksik bilgi"
        Exit Sub
    End If

    On Error GoTo Hata
    hedefDosya = OlayDosyaYolu(no)
    modDosyaIO.TamKayitYaz hedefDosya, _
        modDosyaIO.KayitMetniUret(sozluk, OlayAlanlari())

    modKonsolide.OnerileriYenile
    OneriyiAc no                                  ' gecmis listesi tazelensin

    modUI.BantYaz ws, "dg_bant", _
        "✓  Değerlendirme kaydedildi  ·  " & Format$(Now, "dd.mm.yyyy hh:nn"), _
        modTasarim.CLR_ONAY_ZEMIN, modTasarim.CLR_ONAY_YAZI
    Exit Sub

Hata:
    modUI.Hata "Değerlendirme kaydedilemedi." & vbCrLf & vbCrLf & Err.Description, _
               "Kayıt hatası"
End Sub


' ###########################################################################
'  EKRAN DOLDURMA
' ###########################################################################

' Bir oneriyi degerlendirme ekranina yukler.
'
' Bu yordam bir dugmeden cagrilir; ham bir VBA hata penceresi kullaniciya
' hicbir sey anlatmaz (otomasyonda ise gorunmeden kilitlenir). Bu yuzden
' butun hatalar yakalanip anlasilir bir mesaja cevrilir.
Public Sub OneriyiAc(ByVal oneriNo As String)
    Dim ws As Object, veri As Variant, i As Long
    Dim hataMetni As String

    On Error GoTo Hata

    veri = modKonsolide.VeriDizisi()
    i = modKonsolide.SatirBul(veri, oneriNo)
    If i = 0 Then
        modUI.Hata "Öneri bulunamadı: " & oneriNo, "Değerlendirme"
        Exit Sub
    End If

    Set ws = DegerlendirmeSayfasi()
    modUI.SayfaGoster modUI.SAYFA_DEGERLENDIRME, etkinlestir:=False
    modUI.KorumaKapa ws

    ws.Range("dg_oneri_no").Value = veri(i, modKonsolide.V_ONERI_NO)
    ' Kayitta tarih ISO bicimindedir (sıralanabilir olsun diye); ekranda
    ' okunur bicimde gosterilir.
    ws.Range("dg_tarih").Value = OlayTarihiGoster(CStr(veri(i, modKonsolide.V_TARIH)))
    ws.Range("dg_gonderen").Value = veri(i, modKonsolide.V_AD_SOYAD) & _
                                    "  (" & veri(i, modKonsolide.V_SICIL_NO) & ")"
    ws.Range("dg_baslik").Value = veri(i, modKonsolide.V_BASLIK)
    ws.Range("dg_mevcut").Value = veri(i, modKonsolide.V_MEVCUT)
    ws.Range("dg_cozum").Value = veri(i, modKonsolide.V_COZUM)
    ws.Range("dg_fayda").Value = veri(i, modKonsolide.V_FAYDA)

    ' Giris alanlari son bilinen degerle acilir; ekip yalnizca degistirdigini
    ' duzeltir, her seferinde bastan doldurmaz.
    ws.Range("dg_yeni_durum").Value = veri(i, modKonsolide.V_DURUM)
    ws.Range("dg_not").Value = ""
    modUI.KorumaAc ws

    GecmisiYaz ws, oneriNo
    modUI.BantTemizle ws, "dg_bant"

    ' Imleci ilk giris alanina goturmek gorsel bir kolayliktir; penceresi
    ' olmayan bir oturumda basarisiz olabilir ve bu onemli degildir.
    On Error Resume Next
    ws.Activate
    ws.Range("dg_yeni_durum").Select
    On Error GoTo 0
    Exit Sub

Hata:
    ' Aciklama hemen alinir: sonraki On Error yonergeleri Err'i temizler.
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    On Error Resume Next
    modUI.KorumaAc DegerlendirmeSayfasi()
    On Error GoTo 0
    modUI.Hata "Öneri değerlendirme ekranına yüklenemedi." & vbCrLf & vbCrLf & _
               hataMetni, "Değerlendirme"
End Sub

Private Sub GecmisiYaz(ByVal ws As Object, ByVal oneriNo As String)
    Dim kayit As Object, toplam As Collection
    Dim satir As Long, tablo() As Variant, n As Long, i As Long

    Set toplam = OlayDosyalari(oneriNo)

    modUI.KorumaKapa ws
    ws.Range(ws.Cells(GECMIS_ILK_SATIR, GECMIS_ILK_SUTUN), _
             ws.Cells(GECMIS_ILK_SATIR + GECMIS_AZAMI - 1, GECMIS_SON_SUTUN)).ClearContents

    n = toplam.Count
    If n = 0 Then
        ws.Cells(GECMIS_ILK_SATIR, GECMIS_ILK_SUTUN).Value = _
            "Henüz değerlendirme yapılmamış."
        modUI.KorumaAc ws
        Exit Sub
    End If

    ' Son olaylar ustte gorunsun; ekrana sigmayan eski olaylar dosyada durur.
    ReDim tablo(1 To WorksheetFunction.Min(n, GECMIS_AZAMI), 1 To 4)
    satir = 0
    For i = n To 1 Step -1
        If satir >= GECMIS_AZAMI Then Exit For
        Set kayit = modDosyaIO.KayitOku(CStr(toplam(i)))
        If Not modDosyaIO.KayitTamMi(kayit) Then GoTo SonrakiOlay

        satir = satir + 1
        tablo(satir, 1) = OlayTarihiGoster(modDosyaIO.Al(kayit, "olay_tarihi"))
        tablo(satir, 2) = modDosyaIO.Al(kayit, "degerlendiren_kullanici")
        tablo(satir, 3) = modDosyaIO.Al(kayit, "yeni_durum")
        tablo(satir, 4) = modDosyaIO.Al(kayit, "karar_notu")
SonrakiOlay:
    Next i

    If satir = 0 Then
        ws.Cells(GECMIS_ILK_SATIR, GECMIS_ILK_SUTUN).Value = _
            "Henüz değerlendirme yapılmamış."
        modUI.KorumaAc ws
        Exit Sub
    End If

    ' Not sutunu birlesik oldugu icin blok yazma reddedilir; sutun sutun yazilir.
    Dim r As Long, c As Long
    For r = 1 To satir
        For c = 1 To 4
            ws.Cells(GECMIS_ILK_SATIR + r - 1, GECMIS_ILK_SUTUN + c - 1).Value = tablo(r, c)
        Next c
    Next r

    If n > GECMIS_AZAMI Then
        ws.Cells(GECMIS_ILK_SATIR + GECMIS_AZAMI - 1, GECMIS_SON_SUTUN).Value = _
            "… toplam " & n & " kayıt"
    End If
    modUI.KorumaAc ws
End Sub

Private Function OlayTarihiGoster(ByVal isoTarih As String) As String
    If Len(isoTarih) >= 16 Then
        OlayTarihiGoster = Mid$(isoTarih, 9, 2) & "." & Mid$(isoTarih, 6, 2) & "." & _
                           Left$(isoTarih, 4) & "  " & Mid$(isoTarih, 12, 5)
    Else
        OlayTarihiGoster = isoTarih
    End If
End Function


' ###########################################################################
'  EKRAN -> OLAY KAYDI
' ###########################################################################

Private Function EkrandanOku(ByVal ws As Object, ByVal oneriNo As String) As Object
    Dim sozluk As Object

    Set sozluk = CreateObject("Scripting.Dictionary")
    sozluk.CompareMode = 1

    sozluk("sema") = modAyar.SEMA_SURUMU
    sozluk("oneri_no") = oneriNo
    sozluk("olay_tarihi") = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss")
    sozluk("yeni_durum") = Trim$(CStr(ws.Range("dg_yeni_durum").Value & ""))
    sozluk("karar_notu") = Trim$(CStr(ws.Range("dg_not").Value & ""))
    sozluk("degerlendiren_kullanici") = Environ$("USERNAME")
    sozluk("degerlendiren_bilgisayar") = Environ$("COMPUTERNAME")

    Set EkrandanOku = sozluk
End Function

Private Function Dogrula(ByVal sozluk As Object) As String
    Dim durum As String

    durum = modDosyaIO.Al(sozluk, "yeni_durum")
    If Len(durum) = 0 Then
        Dogrula = "Yeni durum seçilmelidir."
        Exit Function
    End If
    If Not modModel.DurumGecerliMi(durum) Then
        Dogrula = "Geçersiz durum: " & durum
        Exit Function
    End If

    If StrComp(durum, modModel.DURUM_REDDEDILDI, vbTextCompare) = 0 Then
        If Len(modDosyaIO.Al(sozluk, "karar_notu")) < 5 Then
            Dogrula = "Reddedilen bir öneri için karar notu (gerekçe) zorunludur."
            Exit Function
        End If
    End If

    Dogrula = ""
End Function

Private Function OlayDosyaYolu(ByVal oneriNo As String) As String
    Dim klasor As String, ad As String, deneme As Long

    klasor = modAyar.DegerlendirmeYilKlasor(Year(Now))
    modDosyaIO.KlasorZinciriOlustur klasor

    For deneme = 1 To 50
        ad = oneriNo & "_" & modDosyaIO.ZamanDamgasi() & "_" & _
             modDosyaIO.RastgeleOnaltilik(4) & modAyar.UZANTI_KAYIT
        If Not modDosyaIO.DosyaVarMi(klasor & "\" & ad) Then
            OlayDosyaYolu = klasor & "\" & ad
            Exit Function
        End If
    Next deneme

    Err.Raise vbObjectError + 930, "modDegerlendirme.OlayDosyaYolu", _
              "Benzersiz bir olay dosyası adı üretilemedi."
End Function

' Bir onerinin butun olay dosyalarini zaman sirasiyla dondurur.
Public Function OlayDosyalari(ByVal oneriNo As String) As Collection
    Dim tumu As Collection, sonuc As New Collection, yol As Variant
    Dim ad As String, onek As String

    onek = oneriNo & "_"
    Set tumu = modDosyaIO.DosyalariTara(modAyar.DegerlendirmeKlasor(), modAyar.UZANTI_KAYIT)
    For Each yol In tumu
        ad = modDosyaIO.DosyaAdi(CStr(yol))
        If Len(ad) > Len(onek) Then
            If StrComp(Left$(ad, Len(onek)), onek, vbTextCompare) = 0 Then
                sonuc.Add CStr(yol)
            End If
        End If
    Next yol

    Set OlayDosyalari = sonuc
End Function


' ###########################################################################
'  YARDIMCILAR
' ###########################################################################

Public Function DegerlendirmeSayfasi() As Object
    Set DegerlendirmeSayfasi = ThisWorkbook.Worksheets(modUI.SAYFA_DEGERLENDIRME)
End Function


' ###########################################################################
'  TESTLER ICIN -- ekran olmadan degerlendirme
' ###########################################################################
Public Function TestDegerlendirmesi(ByVal oneriNo As String, ByVal durum As String, _
                                    ByVal notu As String) As String
    Dim sozluk As Object, hedefDosya As String

    Set sozluk = CreateObject("Scripting.Dictionary")
    sozluk.CompareMode = 1

    sozluk("sema") = modAyar.SEMA_SURUMU
    sozluk("oneri_no") = oneriNo
    sozluk("olay_tarihi") = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss")
    sozluk("yeni_durum") = durum
    sozluk("karar_notu") = notu
    sozluk("degerlendiren_kullanici") = Environ$("USERNAME")
    sozluk("degerlendiren_bilgisayar") = Environ$("COMPUTERNAME")

    hedefDosya = OlayDosyaYolu(oneriNo)
    modDosyaIO.TamKayitYaz hedefDosya, modDosyaIO.KayitMetniUret(sozluk, OlayAlanlari())
    TestDegerlendirmesi = hedefDosya
End Function
