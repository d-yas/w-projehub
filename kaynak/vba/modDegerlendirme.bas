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
Private Const GECMIS_SON_SUTUN As Long = 6     ' F

Private Function OlayAlanlari() As Variant
    OlayAlanlari = Array("sema", "oneri_no", "olay_tarihi", _
                         "yeni_durum", _
                         "etki_puani", "efor_puani", _
                         "yillik_saat_kazanimi", "yillik_tl_tasarrufu", _
                         "karar_notu", _
                         "degerlendiren_kullanici", "degerlendiren_bilgisayar")
End Function


' ###########################################################################
'  DUGME EYLEMLERI
' ###########################################################################

' Konsoldaki secili satiri degerlendirme ekraninda acar.
Public Sub SeciliyiDegerlendir()
    Dim ws As Object, no As String

    Set ws = modKonsolide.KonsolSayfasi()
    If StrComp(ActiveSheet.Name, ws.Name, vbTextCompare) <> 0 Then
        modUI.SayfaGoster modUI.SAYFA_KONSOL
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

Public Sub KonsolaDon()
    modUI.SayfaGoster modUI.SAYFA_KONSOL
End Sub

' Degerlendirmeyi yeni bir olay dosyasi olarak kaydeder.
Public Sub DegerlendirmeKaydet()
    Dim ws As Object, sozluk As Object
    Dim no As String, hataMetni As String, hedefDosya As String

    Set ws = DegerlendirmeSayfasi()
    no = Trim$(CStr(ws.Range("dg_oneri_no").Value & ""))

    If Len(no) = 0 Then
        modUI.Hata "Ekranda açık bir öneri yok." & vbCrLf & vbCrLf & _
                   "Konsoldan bir öneri seçip “Seçiliyi Değerlendir” " & _
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
    ws.Range("dg_etki").Value = modKonsolide.BosDegilse(veri(i, modKonsolide.V_ETKI))
    ws.Range("dg_efor").Value = modKonsolide.BosDegilse(veri(i, modKonsolide.V_EFOR))
    ws.Range("dg_saat").Value = modKonsolide.BosDegilse(veri(i, modKonsolide.V_SAAT))
    ws.Range("dg_tl").Value = modKonsolide.BosDegilse(veri(i, modKonsolide.V_TL))
    ws.Range("dg_not").Value = ""
    modUI.KorumaAc ws

    OncelikGoster ws                              ' korumasini kendi yonetir
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

' Etki/efor puanlarindan oncelik sinifini gosterir. Formul yerine VBA
' kullanilir: sinif kurali modModel'de tek yerde durur, ekranda kopyasi olmaz.
'
' Oncelik hucresi KILITLIDIR (hesaplanan bir degerdir, elle girilmez), bu
' yuzden yazmadan once koruma acilir. Sayfa korumasini bu yordam kendi yonetir;
' disaridan sarmalanmasi gerekmez.
Public Sub OncelikGoster(Optional ByVal ws As Object = Nothing)
    Dim s As String, etki As Long, efor As Long, hucre As Range

    If ws Is Nothing Then Set ws = DegerlendirmeSayfasi()
    etki = SayiOku(ws.Range("dg_etki").Value)
    efor = SayiOku(ws.Range("dg_efor").Value)
    s = modModel.OncelikSinifi(etki, efor)

    modUI.KorumaKapa ws
    Set hucre = ws.Range("dg_oncelik")
    If Len(s) = 0 Then
        hucre.Value = "— puanlanmadı —"
        hucre.Interior.Color = modTasarim.HexRGB(modTasarim.CLR_BUZ_ZEMIN)
        hucre.Font.Color = modTasarim.HexRGB(modTasarim.CLR_METIN_GRI)
        hucre.Font.Bold = False
        hucre.HorizontalAlignment = xlCenter
    Else
        hucre.Value = s
        modTasarim.OncelikRozetiUygula hucre, s
    End If
    modUI.KorumaAc ws
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
    ReDim tablo(1 To WorksheetFunction.Min(n, GECMIS_AZAMI), 1 To 5)
    satir = 0
    For i = n To 1 Step -1
        If satir >= GECMIS_AZAMI Then Exit For
        Set kayit = modDosyaIO.KayitOku(CStr(toplam(i)))
        If Not modDosyaIO.KayitTamMi(kayit) Then GoTo SonrakiOlay

        satir = satir + 1
        tablo(satir, 1) = OlayTarihiGoster(modDosyaIO.Al(kayit, "olay_tarihi"))
        tablo(satir, 2) = modDosyaIO.Al(kayit, "degerlendiren_kullanici")
        tablo(satir, 3) = modDosyaIO.Al(kayit, "yeni_durum")
        tablo(satir, 4) = PuanGoster(modDosyaIO.Al(kayit, "etki_puani"), _
                                     modDosyaIO.Al(kayit, "efor_puani"))
        tablo(satir, 5) = modDosyaIO.Al(kayit, "karar_notu")
SonrakiOlay:
    Next i

    If satir = 0 Then
        ws.Cells(GECMIS_ILK_SATIR, GECMIS_ILK_SUTUN).Value = _
            "Henüz değerlendirme yapılmamış."
        modUI.KorumaAc ws
        Exit Sub
    End If

    ws.Cells(GECMIS_ILK_SATIR, GECMIS_ILK_SUTUN).Resize(satir, 5).Value = tablo

    If n > GECMIS_AZAMI Then
        ws.Cells(GECMIS_ILK_SATIR + GECMIS_AZAMI - 1, GECMIS_SON_SUTUN).Value = _
            "… toplam " & n & " kayıt"
    End If
    modUI.KorumaAc ws
End Sub

Private Function PuanGoster(ByVal etki As String, ByVal efor As String) As String
    If Len(etki) = 0 And Len(efor) = 0 Then Exit Function
    PuanGoster = "etki " & etki & " / efor " & efor
End Function

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
    sozluk("etki_puani") = SayiMetni(ws.Range("dg_etki").Value)
    sozluk("efor_puani") = SayiMetni(ws.Range("dg_efor").Value)
    sozluk("yillik_saat_kazanimi") = SayiMetni(ws.Range("dg_saat").Value)
    sozluk("yillik_tl_tasarrufu") = SayiMetni(ws.Range("dg_tl").Value)
    sozluk("karar_notu") = Trim$(CStr(ws.Range("dg_not").Value & ""))
    sozluk("degerlendiren_kullanici") = Environ$("USERNAME")
    sozluk("degerlendiren_bilgisayar") = Environ$("COMPUTERNAME")

    Set EkrandanOku = sozluk
End Function

Private Function Dogrula(ByVal sozluk As Object) As String
    Dim durum As String, etki As Long, efor As Long

    durum = modDosyaIO.Al(sozluk, "yeni_durum")
    If Len(durum) = 0 Then
        Dogrula = "Yeni durum seçilmelidir."
        Exit Function
    End If
    If Not modModel.DurumGecerliMi(durum) Then
        Dogrula = "Geçersiz durum: " & durum
        Exit Function
    End If

    etki = SayiOku(modDosyaIO.Al(sozluk, "etki_puani"))
    efor = SayiOku(modDosyaIO.Al(sozluk, "efor_puani"))
    If etki < 0 Or etki > 5 Or efor < 0 Or efor > 5 Then
        Dogrula = "Etki ve efor puanları 1 ile 5 arasında olmalıdır."
        Exit Function
    End If

    ' Kabul edilen bir oneri puanlanmadan ilerlerse oncelik matrisi ve pano
    ' bos kalir; bu asamada puanlama zorunlu tutulur.
    If modModel.KabulEdildiMi(durum) Then
        If etki = 0 Or efor = 0 Then
            Dogrula = "“" & durum & "” durumuna geçmeden önce etki ve efor " & _
                      "puanlarını (1–5) girin."
            Exit Function
        End If
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

Private Function SayiOku(ByVal v As Variant) As Long
    If IsNumeric(v) Then SayiOku = CLng(Val(v)) Else SayiOku = 0
End Function

Private Function SayiMetni(ByVal v As Variant) As String
    If IsNumeric(v) Then
        If Val(v) = 0 Then
            SayiMetni = ""
        Else
            ' Ondalik ayraci makine ayarindan bagimsiz olsun: dosyada nokta.
            SayiMetni = Replace(CStr(CDbl(v)), ",", ".")
        End If
    Else
        SayiMetni = ""
    End If
End Function


' ###########################################################################
'  TESTLER ICIN -- ekran olmadan degerlendirme
' ###########################################################################
Public Function TestDegerlendirmesi(ByVal oneriNo As String, ByVal durum As String, _
                                    ByVal etki As Long, ByVal efor As Long, _
                                    ByVal saat As Double, ByVal tl As Double, _
                                    ByVal notu As String) As String
    Dim sozluk As Object, hedefDosya As String

    Set sozluk = CreateObject("Scripting.Dictionary")
    sozluk.CompareMode = 1

    sozluk("sema") = modAyar.SEMA_SURUMU
    sozluk("oneri_no") = oneriNo
    sozluk("olay_tarihi") = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss")
    sozluk("yeni_durum") = durum
    sozluk("etki_puani") = IIf(etki = 0, "", CStr(etki))
    sozluk("efor_puani") = IIf(efor = 0, "", CStr(efor))
    sozluk("yillik_saat_kazanimi") = IIf(saat = 0, "", Replace(CStr(saat), ",", "."))
    sozluk("yillik_tl_tasarrufu") = IIf(tl = 0, "", Replace(CStr(tl), ",", "."))
    sozluk("karar_notu") = notu
    sozluk("degerlendiren_kullanici") = Environ$("USERNAME")
    sozluk("degerlendiren_bilgisayar") = Environ$("COMPUTERNAME")

    hedefDosya = OlayDosyaYolu(oneriNo)
    modDosyaIO.TamKayitYaz hedefDosya, modDosyaIO.KayitMetniUret(sozluk, OlayAlanlari())
    TestDegerlendirmesi = hedefDosya
End Function
