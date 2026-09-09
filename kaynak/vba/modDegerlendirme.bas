Attribute VB_Name = "modDegerlendirme"
Option Explicit

' NOT: Modul duzeyi bildirimler VBA'da YALNIZCA burada, ilk yordamdan once
' bulunabilir (test_uretim.py bunu denetler).

Private Const GECMIS_ILK_SATIR As Long = 34
Private Const GECMIS_AZAMI As Long = 14
Private Const GECMIS_ILK_SUTUN As Long = 2     ' B
' F dahil: not sutunu E:F birlesiktir ve ClearContents birlesimin yalnizca
' bir parcasi uzerinde calistirilamaz.
Private Const GECMIS_SON_SUTUN As Long = 6     ' F

' ============================================================================
'  modDegerlendirme -- degerlendirme ekrani ve EKLE-ONLY kayit
'
'  Hicbir satir guncellenmez, hicbir satir silinmez. Her degerlendirme
'  "Olaylar" sayfasinin SONUNA yeni bir satir ekler. Bir onerinin gecmisi o
'  satirlarin toplamidir; "kim, ne zaman, neyi degistirdi" izi yapisal olarak
'  korunur -- silinebilecek bir gecmis alani yoktur.
'
'  Sira zaman siradir: satirlar yalnizca sona eklenir ve ekleme, dosya kilidi
'  elde tutulurken yapilir. Eskiden bu guvence dosya adlarindaki 10 ms
'  cozunurluklu zaman damgasindan geliyordu; zaman damgasina artik gerek yok,
'  cunku ayni anda iki kisi yazamaz.
' ============================================================================


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

' ---------------------------------------------------------------------------
'  DegerlendirmeKaydet -- olayi depoya yazar, ekrani tazeler.
'
'  Yazma ve ardindan gelen "cek" TEK bir gizli Excel ornegiyle yapilir
'  (modDepo.OlayEkleVeCek): ayri ayri yapilsaydi iki ornek baslatilir ve
'  kaydetme bir saniye daha uzardi.
' ---------------------------------------------------------------------------
Public Sub DegerlendirmeKaydet()
    Dim ws As Object, sozluk As Object
    Dim no As String, hataMetni As String

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
    modUI.HizliModAc
    DurumCubugu "Değerlendirme kaydediliyor..."

    modDepo.OlayEkleVeCek sozluk, ThisWorkbook
    modKonsolide.OzetYaz modKonsolide.YerelYenile()

    DurumCubugu ""
    modUI.HizliModKapa

    OneriyiAc no                                  ' gecmis listesi tazelensin

    modUI.BantYaz ws, "dg_bant", _
        "✓  Değerlendirme kaydedildi  ·  " & Format$(Now, "dd.mm.yyyy hh:nn"), _
        modTasarim.CLR_ONAY_ZEMIN, modTasarim.CLR_ONAY_YAZI
    Exit Sub

Hata:
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    DurumCubugu ""
    modUI.HizliModKapa

    ' Kuyruk NEDENE gore secilir. "Birkaç saniye sonra yeniden deneyin"
    ' demek, sorun kurulum konumuysa yaniltici olur: orada bekleyerek hicbir
    ' sey duzelmez ve kullanici defalarca dener.
    If modDepo.DepoOneDriveAltindaMi() Then
        modUI.Hata "Değerlendirme kaydedilemedi." & vbCrLf & vbCrLf & _
                   modDepo.OneDriveAciklamasi() & vbCrLf & vbCrLf & _
                   "Yazdıklarınız ekranda duruyor.", "Kayıt hatası"
    Else
        modUI.Hata "Değerlendirme kaydedilemedi." & vbCrLf & vbCrLf & hataMetni & _
                   vbCrLf & vbCrLf & _
                   "Yönetim kitabına o anda başka biri yazıyor olabilir; " & _
                   "birkaç saniye sonra yeniden deneyin. Yazdıklarınız " & _
                   "ekranda duruyor.", "Kayıt hatası"
    End If
End Sub

' Bir hata isleyicisi icinden de cagrildigi icin tamamen hataya dayanikli.
Private Sub DurumCubugu(ByVal metin As String)
    On Error Resume Next
    If Len(metin) = 0 Then
        Application.StatusBar = False
    Else
        Application.StatusBar = metin
    End If
    Err.Clear
    On Error GoTo 0
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
    Dim satirlar As Collection
    Dim olaylar As Variant
    Dim tablo() As Variant
    Dim n As Long, i As Long, r As Long, satir As Long
    Dim c As Long

    olaylar = modDepo.TabloOku(modKonsolide.OlaylarSayfasi(), modDepo.E_SUTUN_SAYISI)
    Set satirlar = OlaySatirlari(olaylar, oneriNo)

    modUI.KorumaKapa ws
    ws.Range(ws.Cells(GECMIS_ILK_SATIR, GECMIS_ILK_SUTUN), _
             ws.Cells(GECMIS_ILK_SATIR + GECMIS_AZAMI - 1, GECMIS_SON_SUTUN)).ClearContents

    n = satirlar.Count
    If n = 0 Then
        ws.Cells(GECMIS_ILK_SATIR, GECMIS_ILK_SUTUN).Value = _
            "Henüz değerlendirme yapılmamış."
        modUI.KorumaAc ws
        Exit Sub
    End If

    ' Son olaylar ustte gorunsun; ekrana sigmayan eski olaylar sayfada durur.
    ReDim tablo(1 To WorksheetFunction.Min(n, GECMIS_AZAMI), 1 To 4)
    satir = 0
    For i = n To 1 Step -1
        If satir >= GECMIS_AZAMI Then Exit For
        r = satirlar(i)
        satir = satir + 1
        tablo(satir, 1) = OlayTarihiGoster(CStr(olaylar(r, modDepo.E_OLAY_TARIHI) & ""))
        tablo(satir, 2) = olaylar(r, modDepo.E_DEGERLENDIREN_KULLANICI)
        tablo(satir, 3) = olaylar(r, modDepo.E_YENI_DURUM)
        tablo(satir, 4) = olaylar(r, modDepo.E_KARAR_NOTU)
    Next i

    ' Not sutunu birlesik oldugu icin blok yazma reddedilir; hucre hucre yazilir.
    For i = 1 To satir
        For c = 1 To 4
            ws.Cells(GECMIS_ILK_SATIR + i - 1, GECMIS_ILK_SUTUN + c - 1).Value = tablo(i, c)
        Next c
    Next i

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

' Bir onerinin olay satirlarinin numaralarini SATIR SIRASIYLA dondurur.
' Satir sirasi zaman sirasidir: satirlar yalnizca sona eklenir.
Public Function OlaySatirlari(ByVal olaylar As Variant, _
                              ByVal oneriNo As String) As Collection
    Dim sonuc As New Collection
    Dim r As Long

    Set OlaySatirlari = sonuc
    If IsEmpty(olaylar) Then Exit Function

    For r = LBound(olaylar, 1) To UBound(olaylar, 1)
        If StrComp(Trim$(CStr(olaylar(r, modDepo.E_ONERI_NO) & "")), _
                   oneriNo, vbTextCompare) = 0 Then
            sonuc.Add r
        End If
    Next r
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


' ###########################################################################
'  YARDIMCILAR
' ###########################################################################

Public Function DegerlendirmeSayfasi() As Object
    Set DegerlendirmeSayfasi = ThisWorkbook.Worksheets(modUI.SAYFA_DEGERLENDIRME)
End Function


' ###########################################################################
'  TESTLER ICIN -- ekran olmadan degerlendirme
'
'  Hata durumunda "HATA: ..." dondurur. Yakalanmamis bir VBA hatasi, gorunmez
'  bir Excel'de gorunmeyen bir pencere acar ve testi sonsuza kadar bekletir;
'  hatayi donen degere cevirmek onu gorulur bir basarisizliga dondurur.
' ###########################################################################
Public Function TestDegerlendirmesi(ByVal oneriNo As String, ByVal durum As String, _
                                    ByVal notu As String) As String
    Dim sozluk As Object

    On Error GoTo Hata

    Set sozluk = CreateObject("Scripting.Dictionary")
    sozluk.CompareMode = 1

    sozluk("sema") = modAyar.SEMA_SURUMU
    sozluk("oneri_no") = oneriNo
    sozluk("olay_tarihi") = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss")
    sozluk("yeni_durum") = durum
    sozluk("karar_notu") = notu
    sozluk("degerlendiren_kullanici") = Environ$("USERNAME")
    sozluk("degerlendiren_bilgisayar") = Environ$("COMPUTERNAME")

    TestDegerlendirmesi = "satir=" & modDepo.OlayEkle(sozluk)
    Exit Function

Hata:
    TestDegerlendirmesi = "HATA: " & Err.Number & " " & Err.Description
End Function

' Bir onerinin bu kitaptaki olay sayisi (once cekilmis olmalidir).
Public Function TestOlaySayisi(ByVal oneriNo As String) As Long
    Dim olaylar As Variant
    olaylar = modDepo.TabloOku(modKonsolide.OlaylarSayfasi(), modDepo.E_SUTUN_SAYISI)
    TestOlaySayisi = OlaySatirlari(olaylar, oneriNo).Count
End Function
