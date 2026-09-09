Attribute VB_Name = "modAyar"
Option Explicit

' ============================================================================
'  modAyar -- dosya yollari, ekran sifreleri, depo parolasi, sema surumu
'
'  EKRAN sifreleri (SIFRE_PERSONEL, SIFRE_YONETIM) yalnizca "yanlis ekrana
'  yanlislikla girmeyi" onler. SIFRE_KORUMA kazara hucre bozmayi onler.
'
'  SIFRE_DOSYA baskadir: yonetim kitabinin ACILIS parolasidir ve artik
'  gizliligin ASIL siniridir. Personel dosyayi kopyalayabilir ama parolasiz
'  acamaz. Tehdit modeli SIRADAN PERSONELDIR: parola bu kaynakta ve uretilen
'  kitaplarin VBA'sinda duz durur, yani VBA'yi acmayi bilen biri onu okur.
'  Bu bilincli bir sinirdir (bkz. TASARIM-VE-GEREKCE.md madde 8).
' ============================================================================

Public Const SEMA_SURUMU As String = "4"

' Oneri numarasinin oneki: PRJ-2026-0001
Public Const ONEK_ONERI_NO As String = "PRJ"

' --- Ekran sifreleri (gercek guvenlik siniri degil) -----------------------
' Birim adi burada degil: kaynak/tasarim.py -> BIRIM_ADI
Public Const SIFRE_PERSONEL As String = "proje"
Public Const SIFRE_YONETIM As String = "proje-yonetim"

' --- Sayfa/kitap koruma parolasi (kazara bozmayi onler) -------------------
Public Const SIFRE_KORUMA As String = "po-koruma"

' --- Yonetim kitabinin acilis parolasi (veri deposunun kapisi) ------------
Public Const SIFRE_DOSYA As String = "proje-depo"

' --- Klasor ve dosya adlari ----------------------------------------------
Public Const KLASOR_YONETIM As String = "yonetim"
Public Const KLASOR_YEDEK As String = "yedek"
Public Const DOSYA_YONETIM As String = "ProjeYonetim.xlsm"

' Kac gunluk yedek kopya saklanir (yenisi eklenince en eskisi silinir).
Public Const YEDEK_ADET As Long = 7


' ---------------------------------------------------------------------------
'  KokKlasor -- ortak "projeoneri\" klasorunun disk yolunu bulur.
'
'  Kitap iki yerden birinde olabilir:
'    projeoneri\ProjeOneri.xlsm            -> kok = kitabin klasoru
'    projeoneri\yonetim\ProjeYonetim.xlsm  -> kok = bir ust klasor
'
'  Karar kitap adina degil klasor yapisina bakilarak verilir: "yonetim" alt
'  klasoru hangi seviyede varsa kok orasidir.
'
'  Isaret olarak "yonetim" kullanilir: personel bu klasoru gorebilir, icindeki
'  kitap ise parolalidir.
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
'  YonetimKitapYolu -- VERI DEPOSUNUN yolu.
'
'  Ayri kayit dosyalari ve yil klasorleri YOKTUR. Butun veri bu kitabin
'  icindeki iki cok gizli sayfada durur: "Oneriler" ve "Olaylar"
'  (bkz. modDepo). Sistemde uretilen tek diger dosya gunluk yedek kopyadir.
'
'  Kitabin kendisi calisiyorsa kendi tam yolu kullanilir: dosya yeniden
'  adlandirilmis olabilir ve depo yine kendisidir.
' ---------------------------------------------------------------------------
Public Function YonetimKitapYolu() As String
    Dim kendi As String

    kendi = modDosyaIO.YolTemizle(modDosyaIO.YerelYol(ThisWorkbook.Path))
    If StrComp(modDosyaIO.KlasorAdi(kendi), KLASOR_YONETIM, vbTextCompare) = 0 Then
        YonetimKitapYolu = modDosyaIO.YerelYol(ThisWorkbook.FullName)
        Exit Function
    End If

    YonetimKitapYolu = YonetimKlasor() & "\" & DOSYA_YONETIM
End Function

Public Function YedekKlasor() As String
    YedekKlasor = YonetimKlasor() & "\" & KLASOR_YEDEK
End Function
