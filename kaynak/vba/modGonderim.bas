Attribute VB_Name = "modGonderim"
Option Explicit

' NOT: Modul duzeyi bildirimler (Const, Dim) VBA'da YALNIZCA burada, ilk
' yordamdan once bulunabilir. Yordamlar arasina konan bir Const derleme
' hatasi verir; uustelik VBA modulleri yordam bazinda derledigi icin hata
' yalnizca o yordam ilk cagrildiginda ortaya cikar.

' Ayni oneri numarasina denk gelinirse kac kez yeniden denenecegi.
Private Const AZAMI_DENEME As Long = 12

' ============================================================================
'  modGonderim -- personel tarafi: giris, form dogrulama, gonderim
'
'  Sistemin temel kurali burada uygulanir: KIMSE ORTAK BIR DOSYAYA YAZMAZ.
'  Her gonderim "yonetim\oneriler\<yil>\" altina KENDI dosyasini birakir. Iki kisi ayni
'  anda gonderse bile ayri dosyalara yazarlar; cakisma onlenmis degil,
'  yapisal olarak imkansizdir.
'
'  Form alanlari hucre adresiyle degil ADLANDIRILMIS ARALIKLARLA baglanir
'  (frm_ad_soyad gibi). Boylece sayfa duzeni degistiginde kod degismez.
' ============================================================================

' Kayit dosyasindaki alan sirasi. Dosyayi Not Defteri'yle acan biri icin
' okunabilir bir sira: once kimlik, sonra sorun, sonra oneri, sonra iz kaydi.
Private Function KayitAlanlari() As Variant
    KayitAlanlari = Array("sema", "oneri_no", "tarih", _
                          "ad_soyad", "sicil_no", _
                          "mevcut_durum", _
                          "oneri_basligi", "cozum_onerisi", "beklenen_fayda", _
                          "gonderen_bilgisayar", "gonderen_kullanici")
End Function


' ###########################################################################
'  DUGME EYLEMLERI
' ###########################################################################

' "Sisteme Gir" dugmesi -- Giris sayfasi
Public Sub SistemeGir()
    If Not modUI.SifreDogrula(modAyar.SIFRE_PERSONEL, "Proje Öneri Formu") Then Exit Sub
    modUI.OturumAc Array(modUI.SAYFA_FORM), modUI.SAYFA_FORM
    modUI.BantTemizle FormSayfasi(), "frm_bant"
End Sub

' "Çıkış" dugmesi -- form sayfasi
Public Sub Cikis()
    modUI.OturumKapat
End Sub

' "Gönder" dugmesi -- form sayfasi
Public Sub OneriGonder()
    Dim ws As Object
    Dim sozluk As Object
    Dim hataMetni As String
    Dim oneriNo As String

    Set ws = FormSayfasi()
    modUI.BantTemizle ws, "frm_bant"

    Set sozluk = FormdanOku(ws)
    hataMetni = Dogrula(sozluk)
    If Len(hataMetni) > 0 Then
        modUI.BantYaz ws, "frm_bant", "⚠  " & hataMetni, _
                      modTasarim.CLR_UYARI_ZEMIN, modTasarim.CLR_UYARI_YAZI
        modUI.Hata hataMetni, "Eksik bilgi"
        Exit Sub
    End If

    On Error GoTo Hata

    oneriNo = KaydiYaz(sozluk)

    FormuTemizle
    modUI.BantYaz ws, "frm_bant", _
        "✓  Öneriniz alındı.  Öneri numaranız: " & oneriNo, _
        modTasarim.CLR_ONAY_ZEMIN, modTasarim.CLR_ONAY_YAZI

    modUI.Bilgi "Öneriniz başarıyla gönderildi." & vbCrLf & vbCrLf & _
                "Öneri numaranız:" & vbCrLf & oneriNo & vbCrLf & vbCrLf & _
                "Bu numarayı not almanız, önerinizin durumunu sormak " & _
                "istediğinizde işinizi kolaylaştırır.", _
                "Gönderim tamamlandı"
    Exit Sub

Hata:
    modUI.BantYaz ws, "frm_bant", "⚠  Gönderim yapılamadı.", _
                  modTasarim.CLR_UYARI_ZEMIN, modTasarim.CLR_UYARI_YAZI
    modUI.Hata "Öneri kaydedilemedi." & vbCrLf & vbCrLf & _
               Err.Description & vbCrLf & vbCrLf & _
               "Ortak klasöre erişiminiz olduğundan emin olun; sorun " & _
               "sürerse Değerlendirme ekibine başvurun.", "Gönderim hatası"
End Sub

' "Temizle" dugmesi -- form sayfasi
Public Sub FormuTemizle()
    Dim ws As Object
    Dim ad As Variant
    Dim hataMetni As String

    Set ws = FormSayfasi()

    On Error GoTo Hata
    modUI.KorumaKapa ws
    For Each ad In FormAlanAdlari()
        ' Alan() birlesik hucrenin tamamini dondurur; parcasi uzerinde
        ' ClearContents calismaz.
        modUI.Alan(ws, CStr(ad)).ClearContents
    Next ad
    modUI.KorumaAc ws

    On Error Resume Next
    ws.Range("frm_ad_soyad").Select
    On Error GoTo 0
    Exit Sub

Hata:
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    On Error Resume Next
    modUI.KorumaAc ws
    On Error GoTo 0
    modUI.Hata "Form temizlenemedi." & vbCrLf & vbCrLf & hataMetni, "Öneri Formu"
End Sub


' ###########################################################################
'  KAYIT YAZMA
' ###########################################################################

' ---------------------------------------------------------------------------
'  KaydiYaz -- numara uret, yaz; ad kapilmissa yeni numarayla yeniden dene.
'
'  Numara kisa oldugu icin (gun + uc karakter) ayni gun icinde ayni numaranin
'  iki kez uretilmesi -- olasi olmasa da -- mumkundur. Yazma denemesi hedef
'  dosya varsa BASARISIZ doner ve hicbir seyin uzerine yazilmaz; burada yeni
'  bir numara uretilip yeniden denenir. Boylece kisa numara, kayit kaybi
'  riski getirmeden kullanilabilir.
' ---------------------------------------------------------------------------
Private Function KaydiYaz(ByVal sozluk As Object) As String
    Dim klasor As String, oneriNo As String, hedefDosya As String
    Dim deneme As Long, aciklama As String

    klasor = modAyar.OnerilerYilKlasor(Year(Now))
    modDosyaIO.KlasorZinciriOlustur klasor

    For deneme = 1 To AZAMI_DENEME
        oneriNo = OneriNoUret()
        hedefDosya = klasor & "\" & oneriNo & modAyar.UZANTI_KAYIT

        sozluk("sema") = modAyar.SEMA_SURUMU
        sozluk("oneri_no") = oneriNo
        sozluk("tarih") = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss")
        sozluk("gonderen_bilgisayar") = Environ$("COMPUTERNAME")
        sozluk("gonderen_kullanici") = Environ$("USERNAME")

        If modDosyaIO.TamKayitYazmayiDene( _
               hedefDosya, modDosyaIO.KayitMetniUret(sozluk, KayitAlanlari()), _
               aciklama) Then
            KaydiYaz = oneriNo
            Exit Function
        End If
    Next deneme

    ' Buraya gelinmesinin gercekci tek nedeni izin sorunudur: ayni gun icinde
    ' art arda on iki kez ayni numaraya denk gelme ihtimali yok denecek kadar
    ' kucuktur.
    Err.Raise vbObjectError + 920, "modGonderim.KaydiYaz", _
              "Öneri klasörüne yazılamadı." & vbCrLf & aciklama & vbCrLf & _
              "Ortak klasöre yazma izniniz olmayabilir."
End Function

' ---------------------------------------------------------------------------
'  Oneri numarasi:  PRJ-26A7K   (PRJ - yil - uc rastgele karakter)
'
'  Telefonda soylenebilmeli, elle yazilabilmeli, bir yere not edilebilmeli.
'
'  Sirali numara (000001) kullanilamaz: sirali sayac ortak bir dosya
'  gerektirir, o da "kimse ortak dosyaya yazmaz" kuralini bozar. Rastgele
'  ek, merkezi bir sayac olmadan benzersizligi saglar; ayni yil icindeki
'  kalan catisma ihtimalini KaydiYaz'daki yeniden deneme kapatir.
'
'  Yil neden duruyor: dosyalar yil klasorlerine yazilir ve "varsa
'  olusturma" kontrolu yalnizca ayni klasorde calisir. Yil olmasaydi iki
'  farkli yilda uretilen ayni numara sessizce iki oneriye verilirdi.
' ---------------------------------------------------------------------------
Public Function OneriNoUret() As String
    OneriNoUret = modAyar.ONEK_ONERI_NO & "-" & Format$(Now, "yy") & _
                  modDosyaIO.KisaRastgele(3)
End Function


' ###########################################################################
'  FORM <-> KAYIT
' ###########################################################################

Private Function FormSayfasi() As Object
    Set FormSayfasi = ThisWorkbook.Worksheets(modUI.SAYFA_FORM)
End Function

' Adlandirilmis aralik adi -> kayit alani adi eslemesi.
Private Function FormAlanAdlari() As Variant
    FormAlanAdlari = Array("frm_ad_soyad", "frm_sicil_no", _
                           "frm_mevcut", "frm_baslik", _
                           "frm_cozum", "frm_fayda")
End Function

Private Function AlanKarsiligi(ByVal aralikAdi As String) As String
    Select Case aralikAdi
        Case "frm_ad_soyad": AlanKarsiligi = "ad_soyad"
        Case "frm_sicil_no": AlanKarsiligi = "sicil_no"
        Case "frm_mevcut":   AlanKarsiligi = "mevcut_durum"
        Case "frm_baslik":   AlanKarsiligi = "oneri_basligi"
        Case "frm_cozum":    AlanKarsiligi = "cozum_onerisi"
        Case "frm_fayda":    AlanKarsiligi = "beklenen_fayda"
        Case Else:           AlanKarsiligi = ""
    End Select
End Function

Private Function FormdanOku(ByVal ws As Object) As Object
    Dim sozluk As Object, ad As Variant, alan As String
    Dim v As Variant

    Set sozluk = CreateObject("Scripting.Dictionary")
    sozluk.CompareMode = 1

    For Each ad In FormAlanAdlari()
        alan = AlanKarsiligi(CStr(ad))
        v = ""
        On Error Resume Next
        v = ws.Range(CStr(ad)).Cells(1, 1).Value
        On Error GoTo 0
        sozluk(alan) = Trim$(CStr(v & ""))
    Next ad

    ' Yazilacak ama formda olmayan alanlar
    sozluk("sema") = modAyar.SEMA_SURUMU
    sozluk("oneri_no") = ""
    sozluk("tarih") = ""
    sozluk("gonderen_bilgisayar") = ""
    sozluk("gonderen_kullanici") = ""

    Set FormdanOku = sozluk
End Function


' ###########################################################################
'  DOGRULAMA
' ###########################################################################

' Bos dizge dondurmesi "gecerli" demektir.
Private Function Dogrula(ByVal sozluk As Object) As String
    Dim eksik As String

    If Len(modDosyaIO.Al(sozluk, "ad_soyad")) < 3 Then
        eksik = "Ad Soyad"
    ElseIf Len(modDosyaIO.Al(sozluk, "sicil_no")) = 0 Then
        eksik = "Sicil No"
    ElseIf Len(modDosyaIO.Al(sozluk, "mevcut_durum")) < 10 Then
        eksik = "Mevcut Durum (en az 10 karakter)"
    ElseIf Len(modDosyaIO.Al(sozluk, "oneri_basligi")) < 5 Then
        eksik = "Öneri Başlığı (en az 5 karakter)"
    ElseIf Len(modDosyaIO.Al(sozluk, "cozum_onerisi")) < 10 Then
        eksik = "Çözüm Öneriniz (en az 10 karakter)"
    End If

    If Len(eksik) > 0 Then
        Dogrula = "Lütfen şu alanı doldurun: " & eksik
        Exit Function
    End If

    Dogrula = ""
End Function


' ###########################################################################
'  TESTLER ICIN -- ekran olmadan gonderim
'
'  Testler formu doldurup dugmeye basamaz; bu kapi ayni yazma yolunu
'  kullanarak dogrudan kayit uretir. Uretim kodu ile test kodu ayni
'  KaydiYaz cagrisindan gecer.
' ###########################################################################
' Hata durumunda "HATA: ..." dondurur. Yakalanmamis bir VBA hatasi, gorunmez
' bir Excel'de gorunmeyen bir pencere acar ve testi sonsuza kadar bekletir;
' hatayi donen degere cevirmek onu goruldur bir basarisizliga dondurur.
Public Function TestGonderimi(ByVal adSoyad As String, ByVal sicilNo As String, _
                              ByVal mevcut As String, ByVal baslik As String, _
                              ByVal cozum As String, ByVal fayda As String) As String
    Dim sozluk As Object

    On Error GoTo Hata

    Set sozluk = CreateObject("Scripting.Dictionary")
    sozluk.CompareMode = 1

    sozluk("ad_soyad") = adSoyad
    sozluk("sicil_no") = sicilNo
    sozluk("mevcut_durum") = mevcut
    sozluk("oneri_basligi") = baslik
    sozluk("cozum_onerisi") = cozum
    sozluk("beklenen_fayda") = fayda

    TestGonderimi = KaydiYaz(sozluk)
    Exit Function

Hata:
    TestGonderimi = "HATA: " & Err.Number & " " & Err.Description
End Function
