Attribute VB_Name = "modGonderim"
Option Explicit

' ============================================================================
'  modGonderim -- personel tarafi: giris, form dogrulama, gonderim
'
'  Gonderim, yonetim kitabinin icindeki "Oneriler" sayfasina BIR SATIR ekler.
'  Ayri kayit dosyasi yoktur. Yazmanin nasil catismasiz yapildigi ve neden
'  ayri bir gizli Excel ornegi kullanildigi modDepo'nun basinda anlatilir;
'  bu modul o kapiyi cagirir, kilit ve yeniden deneme oraya aittir.
'
'  Form alanlari hucre adresiyle degil ADLANDIRILMIS ARALIKLARLA baglanir
'  (frm_ad_soyad gibi). Boylece sayfa duzeni degistiginde kod degismez.
' ============================================================================


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
    modUI.Hata "Öneriniz kaydedilemedi." & vbCrLf & vbCrLf & _
               Err.Description & vbCrLf & vbCrLf & _
               "Yönetim kitabı şu anda başka bir kullanıcı tarafından " & _
               "kullanılıyor olabilir. Birkaç saniye sonra yeniden deneyin; " & _
               "yazdıklarınız formda duruyor. Sorun sürerse Değerlendirme " & _
               "ekibine başvurun.", "Gönderim hatası"
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
'  KaydiYaz -- iz kaydi alanlarini doldurup depoya bir satir ekletir.
'
'  Oneri NUMARASI burada uretilmez: numara sirali oldugu icin ancak dosya
'  kilidi elimizdeyken, yani depo yazma kipinde acikken guvenle uretilebilir.
'  Onu modDepo yapar ve uretilen numarayi geri verir.
' ---------------------------------------------------------------------------
Private Function KaydiYaz(ByVal sozluk As Object) As String
    sozluk("sema") = modAyar.SEMA_SURUMU
    sozluk("tarih") = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss")
    sozluk("gonderen_bilgisayar") = Environ$("COMPUTERNAME")
    sozluk("gonderen_kullanici") = Environ$("USERNAME")

    KaydiYaz = modDepo.OneriEkle(sozluk)
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
