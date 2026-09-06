Attribute VB_Name = "modAyar"
Option Explicit

' ============================================================================
'  modAyar -- klasor yollari, ekran sifreleri, sema surumu
'
'  Sifreler burada duz metin olarak durur ve GERCEK bir guvenlik siniri
'  DEGILDIR. Amaclari yalnizca "yanlis ekrana yanlislikla girmeyi" onlemektir.
'  Asil erisim denetimi ag klasorunun NTFS izinleridir (bkz. KURULUM.md).
'  Tek gercek sinir yonetim raporunun AES parolasidir; o da koda gomulmez,
'  rapor uretilirken kullaniciya sorulur.
' ============================================================================

Public Const SEMA_SURUMU As String = "3"

' Oneri numarasinin oneki: ON-260904-A7K
Public Const ONEK_ONERI_NO As String = "ON"

' --- Ekran sifreleri (gercek guvenlik siniri degil) -----------------------
Public Const SIFRE_PERSONEL As String = "kaizen"
Public Const SIFRE_YONETIM As String = "kaizen-yonetim"

' --- Sayfa/kitap koruma parolasi (kazara bozmayi onler) -------------------
Public Const SIFRE_KORUMA As String = "kzn-koruma"

' --- Klasor adlari --------------------------------------------------------
Public Const KLASOR_YONETIM As String = "yonetim"
Public Const KLASOR_ONERILER As String = "oneriler"
Public Const KLASOR_DEGERLENDIRME As String = "degerlendirme"
Public Const KLASOR_RAPOR As String = "rapor"

' --- Dosya bicimi ---------------------------------------------------------
Public Const UZANTI_KAYIT As String = ".txt"

' Kayit dosyasinin SON satiri. Okuyucu bu satiri gormezse dosyayi yarim
' yazilmis sayar ve yok sayar. Yazma sirasinda bir kesinti olursa (ag koptu,
' Excel kapandi) yarim kalan dosya asla gecerli bir kayit gibi okunmaz.
Public Const ALAN_SON As String = "kayit_sonu"

' Cok satirli alanlarda satir sonu yerine yazilan belirtec.
' Boylece her kayit alani dosyada tam olarak bir satir kaplar.
Public Const SATIR_BELIRTEC As String = "<|>"

' --- Kaizen etki/efor esikleri -------------------------------------------
Public Const ESIK_YUKSEK_ETKI As Long = 3     ' etki >= 3 ise yuksek
Public Const ESIK_DUSUK_EFOR As Long = 2      ' efor <= 2 ise dusuk


' ---------------------------------------------------------------------------
'  KokKlasor -- ortak "kaizen\" klasorunun disk yolunu bulur.
'
'  Kitap iki yerden birinde olabilir:
'    kaizen\KaizenOneri.xlsm            -> kok = kitabin klasoru
'    kaizen\yonetim\KaizenYonetim.xlsm  -> kok = bir ust klasor
'
'  Karar kitap adina degil klasor yapisina bakilarak verilir: "yonetim" alt
'  klasoru hangi seviyede varsa kok orasidir. Boylece kitap yeniden
'  adlandirilsa bile sistem calismaya devam eder.
'
'  Isaret olarak "yonetim" kullanilir: personel bu klasorun ICINI goremez ama
'  "kaizen\" klasorunu okuyabildigi icin orada "yonetim" adli bir klasor
'  oldugunu gorebilir. Ic klasorler isaret olarak kullanilamaz.
' ---------------------------------------------------------------------------
Public Function KokKlasor() As String
    Dim kendi As String, ust As String

    kendi = modDosyaIO.YerelYol(ThisWorkbook.Path)
    If Len(kendi) = 0 Then
        Err.Raise vbObjectError + 900, "modAyar.KokKlasor", _
                  "Çalışma kitabının klasörü belirlenemedi."
    End If
    kendi = modDosyaIO.YolTemizle(kendi)

    ' Kitap "yonetim\" icindeyse kok bir ust klasordur.
    If StrComp(modDosyaIO.KlasorAdi(kendi), KLASOR_YONETIM, vbTextCompare) = 0 Then
        ust = modDosyaIO.UstKlasor(kendi)
        If Len(ust) > 0 Then
            KokKlasor = ust
            Exit Function
        End If
    End If

    If modDosyaIO.KlasorVarMi(kendi & "\" & KLASOR_YONETIM) Then
        KokKlasor = kendi
        Exit Function
    End If

    ust = modDosyaIO.UstKlasor(kendi)
    If Len(ust) > 0 Then
        If modDosyaIO.KlasorVarMi(ust & "\" & KLASOR_YONETIM) Then
            KokKlasor = ust
            Exit Function
        End If
    End If

    ' Klasor yapisi henuz kurulmamis: kitabin kendi klasoru kok sayilir.
    KokKlasor = kendi
End Function

Public Function YonetimKlasor() As String
    YonetimKlasor = KokKlasor() & "\" & KLASOR_YONETIM
End Function

' ---------------------------------------------------------------------------
'  Oneriler -- gonderimlerin TEK ve KALICI yeri: "yonetim\oneriler\<yil>\"
'
'  Ayri bir "gelen kutusu" YOKTUR. Personel dogrudan buraya yazar; dosya
'  bastan itibaren kalici yerindedir, sonradan hicbir yere tasinmaz.
'
'  Klasor bir BIRAKMA KUTUSU gibi calisir: personel buraya yazabilir ama
'  icini goremez -- listeleyemez, kimsenin onerisini okuyamaz, hicbir dosyayi
'  silemez (NTFS izinleri; bkz. KURULUM.md). Ustelik "yonetim\" klasoru
'  personele yalnizca GECIS hakki verir, listeleme hakki vermez; yani
'  kaizen\ altinda gorunen tek sey calisma kitabi ve "yonetim" adidir.
' ---------------------------------------------------------------------------
Public Function OnerilerKlasor() As String
    OnerilerKlasor = YonetimKlasor() & "\" & KLASOR_ONERILER
End Function

Public Function DegerlendirmeKlasor() As String
    DegerlendirmeKlasor = YonetimKlasor() & "\" & KLASOR_DEGERLENDIRME
End Function

Public Function RaporKlasor() As String
    RaporKlasor = YonetimKlasor() & "\" & KLASOR_RAPOR
End Function

' ---------------------------------------------------------------------------
'  Yil alt klasorleri: oneriler\2026\, degerlendirme\2026\
'  Klasor basina dosya sayisini dusuk tutar; yillik arsivleme dogal olur.
' ---------------------------------------------------------------------------
Public Function OnerilerYilKlasor(ByVal yil As Long) As String
    OnerilerYilKlasor = OnerilerKlasor() & "\" & CStr(yil)
End Function

Public Function DegerlendirmeYilKlasor(ByVal yil As Long) As String
    DegerlendirmeYilKlasor = DegerlendirmeKlasor() & "\" & CStr(yil)
End Function
