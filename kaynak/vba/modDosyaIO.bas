Attribute VB_Name = "modDosyaIO"
Option Explicit

' ============================================================================
'  modDosyaIO -- yol islemleri
'
'  Bu modul bir zamanlar sistemin veri katmaniydi: her gonderim ve her
'  degerlendirme ayri bir UTF-8 metin dosyasiydi ve buradaki ADODB.Stream
'  yordamlari onlari yazip okurdu. Veri artik yonetim kitabinin ICINDEKI iki
'  sayfada duruyor (bkz. modDepo), yani kayit bicimi, atomik yazma, klasor
'  tarama ve zaman damgasi uretimi tumden gereksiz kaldi ve silindi.
'
'  Geriye tek bir is kaldi ve o hala sart:
'
'  ONEDRIVE YOLU. Klasor OneDrive/SharePoint ile eslenmisse Excel kitabin
'  konumunu "https://..." adresi olarak bildirir; VBA boyle bir adresi
'  kullanamaz. YerelYol() bunu diskteki gercek yola cevirir. UNC yollari
'  (\\sunucu\paylasim\...) dokunulmadan gecer; uretim ortami budur.
' ============================================================================


' Bir klasorde taranacak azami dosya sayisi -- bkz. OneDriveKokleri.
Private Const AZAMI_KLASOR_TARAMA As Long = 200


' ###########################################################################
'  YOL ISLEMLERI
' ###########################################################################

' Sondaki ters bolu isaretlerini temizler: "C:\a\b\" -> "C:\a\b"
Public Function YolTemizle(ByVal yol As String) As String
    Dim s As String
    s = Trim$(yol)
    Do While Len(s) > 3 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    YolTemizle = s
End Function

Public Function UstKlasor(ByVal yol As String) As String
    Dim s As String, p As Long
    s = YolTemizle(yol)
    p = InStrRev(s, "\")
    If p <= 0 Then
        UstKlasor = ""
    ElseIf p <= 3 Then                      ' "C:\" kokunun ustu yoktur
        UstKlasor = Left$(s, p)
    Else
        UstKlasor = Left$(s, p - 1)
    End If
End Function

Public Function KlasorAdi(ByVal yol As String) As String
    Dim s As String, p As Long
    s = YolTemizle(yol)
    p = InStrRev(s, "\")
    If p <= 0 Then KlasorAdi = s Else KlasorAdi = Mid$(s, p + 1)
End Function

Public Function DosyaAdi(ByVal yol As String) As String
    Dim p As Long
    p = InStrRev(yol, "\")
    If p <= 0 Then DosyaAdi = yol Else DosyaAdi = Mid$(yol, p + 1)
End Function

Public Function KlasorVarMi(ByVal yol As String) As Boolean
    On Error Resume Next
    KlasorVarMi = (Len(Dir$(YolTemizle(yol), vbDirectory)) > 0)
    If Err.Number <> 0 Then
        KlasorVarMi = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

Public Function DosyaVarMi(ByVal yol As String) As Boolean
    On Error Resume Next
    DosyaVarMi = (Len(Dir$(yol, vbNormal)) > 0)
    If Err.Number <> 0 Then
        DosyaVarMi = False
        Err.Clear
    End If
    On Error GoTo 0
End Function

' Ic ice klasorleri tek seferde olusturur (mkdir -p karsiligi).
Public Sub KlasorZinciriOlustur(ByVal yol As String)
    Dim s As String, ust As String
    s = YolTemizle(yol)
    If Len(s) = 0 Then Exit Sub
    If KlasorVarMi(s) Then Exit Sub

    ust = UstKlasor(s)
    If Len(ust) > 0 And StrComp(ust, s, vbTextCompare) <> 0 Then
        KlasorZinciriOlustur ust
    End If

    On Error Resume Next
    MkDir s
    On Error GoTo 0
End Sub


' ---------------------------------------------------------------------------
'  YerelYol -- OneDrive/SharePoint adresini diskteki gercek yola cevirir.
'
'  Ornek:  https://kurum.sharepoint.com/personal/x/Documents/Belgeler/projeoneri
'          -> C:\Users\x\OneDrive - Kurum\Belgeler\projeoneri
'
'  Yontem: adresin sonundan baslayarak gitgide uzayan parcalar, bilinen yerel
'  OneDrive koklerinin altinda aranir; diskte gercekten var olan ilk eslesme
'  dondurulur. Eslesme bulunamazsa adres oldugu gibi geri verilir -- cagiran
'  taraf o zaman anlamli bir hata gosterir.
' ---------------------------------------------------------------------------
Public Function YerelYol(ByVal yol As String) As String
    Dim parcalar() As String
    Dim kokler As Collection
    Dim kok As Variant
    Dim i As Long, j As Long
    Dim kuyruk As String, aday As String
    Dim temiz As String

    YerelYol = yol
    If Len(yol) = 0 Then Exit Function
    If InStr(1, yol, "http", vbTextCompare) <> 1 Then Exit Function

    ' Protokol + sunucu adini at, "/" -> "\", %20 gibi kacislari coz.
    temiz = yol
    i = InStr(1, temiz, "://")
    If i > 0 Then temiz = Mid$(temiz, i + 3)
    i = InStr(temiz, "/")
    If i > 0 Then temiz = Mid$(temiz, i + 1) Else temiz = ""
    temiz = Replace(temiz, "/", "\")
    temiz = UrlCoz(temiz)
    If Len(temiz) = 0 Then Exit Function

    parcalar = Split(temiz, "\")

    Set kokler = OneDriveKokleri()
    For Each kok In kokler
        ' En uzun kuyruktan en kisaya: en ozgul eslesme kazanir.
        For i = 0 To UBound(parcalar)
            kuyruk = ""
            For j = i To UBound(parcalar)
                If Len(parcalar(j)) > 0 Then
                    If Len(kuyruk) > 0 Then kuyruk = kuyruk & "\"
                    kuyruk = kuyruk & parcalar(j)
                End If
            Next j
            If Len(kuyruk) > 0 Then
                aday = YolTemizle(CStr(kok)) & "\" & kuyruk
                If KlasorVarMi(aday) Or DosyaVarMi(aday) Then
                    YerelYol = aday
                    Exit Function
                End If
            End If
        Next i
    Next kok
End Function

' Ortam degiskenlerinden ve kullanici profilinden olasi OneDrive koklerini toplar.
Private Function OneDriveKokleri() As Collection
    Dim c As New Collection
    Dim ad As Variant, deger As String
    Dim profil As String, klasorAdi As String
    Dim sayac As Long

    For Each ad In Array("OneDrive", "OneDriveCommercial", "OneDriveConsumer")
        deger = Environ$(CStr(ad))
        If Len(deger) > 0 Then KokEkle c, deger
    Next ad

    ' "OneDrive - <Kurum>" bicimindeki klasorler ortam degiskeninde her zaman
    ' gorunmez; kullanici profilinin altini da tararız.
    ' DIKKAT: Dir$ hata verdiginde deger DONDURMEZ; atama yapilmadigi icin
    ' degisken eski degerinde kalir ve "Do While Len(...) > 0" hep dogru olur.
    ' Klasor okunamadigi anda dongu sonsuza gider. Bu yuzden her cagridan
    ' sonra Err denetlenir ve ayrica bir ust sinir vardir.
    profil = Environ$("USERPROFILE")
    If Len(profil) > 0 Then
        On Error Resume Next
        Err.Clear
        klasorAdi = Dir$(profil & "\OneDrive*", vbDirectory)
        If Err.Number <> 0 Then
            Err.Clear
            klasorAdi = ""
        End If
        sayac = 0
        Do While Len(klasorAdi) > 0 And sayac < AZAMI_KLASOR_TARAMA
            sayac = sayac + 1
            If klasorAdi <> "." And klasorAdi <> ".." Then
                KokEkle c, profil & "\" & klasorAdi
            End If
            Err.Clear
            klasorAdi = Dir$
            If Err.Number <> 0 Then
                Err.Clear
                Exit Do
            End If
        Loop
        On Error GoTo 0
    End If

    Set OneDriveKokleri = c
End Function

Private Sub KokEkle(ByRef c As Collection, ByVal yol As String)
    Dim v As Variant
    Dim temiz As String
    temiz = YolTemizle(yol)
    If Len(temiz) = 0 Then Exit Sub
    If Not KlasorVarMi(temiz) Then Exit Sub
    For Each v In c
        If StrComp(CStr(v), temiz, vbTextCompare) = 0 Then Exit Sub
    Next v
    c.Add temiz
End Sub

' "%20" gibi yuzde kacislarini cozer.
Private Function UrlCoz(ByVal s As String) As String
    Dim sonuc As String, i As Long, ch As String
    i = 1
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = "%" And i + 2 <= Len(s) Then
            On Error Resume Next
            sonuc = sonuc & ChrW$(CLng("&H" & Mid$(s, i + 1, 2)))
            If Err.Number <> 0 Then
                sonuc = sonuc & ch
                Err.Clear
                i = i - 2
            End If
            On Error GoTo 0
            i = i + 3
        Else
            sonuc = sonuc & ch
            i = i + 1
        End If
    Loop
    UrlCoz = sonuc
End Function


' ###########################################################################
'  SOZLUK YARDIMCISI
' ###########################################################################

' Sozlukten guvenli okuma (anahtar yoksa varsayilan doner).
' Gonderim ve degerlendirme dogrulamalari bunu kullanir.
Public Function Al(ByVal sozluk As Object, ByVal anahtar As String, _
                   Optional ByVal varsayilan As String = "") As String
    If sozluk Is Nothing Then
        Al = varsayilan
    ElseIf sozluk.Exists(anahtar) Then
        Al = CStr(sozluk(anahtar))
    Else
        Al = varsayilan
    End If
End Function
