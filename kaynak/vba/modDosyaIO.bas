Attribute VB_Name = "modDosyaIO"
Option Explicit

' ============================================================================
'  modDosyaIO -- sistemin guvenilirlik cekirdegi
'
'  Uc sorunu tek yerde cozer:
'
'  1) KODLAMA. VBA'nin yerlesik Open/Print komutu dosyayi ANSI yazar; Turkce
'     karakterler baska bir makinede bozulur. Butun okuma/yazma ADODB.Stream
'     uzerinden UTF-8 (BOM'suz) yapilir.
'
'  2) YARIM DOSYA. Once ".tmp" yazilir, sonra ".txt" olarak yeniden adlandirilir.
'     Yeniden adlandirma dosya sisteminde atomiktir; yonetim tarafi asla
'     yarim yazilmis bir dosya okumaz.
'
'  3) ONEDRIVE YOLU. Klasor OneDrive/SharePoint ile eslenmisse Excel kitabin
'     konumunu "https://..." adresi olarak bildirir ve VBA boyle bir adrese
'     yazamaz. YerelYol() bunu diskteki gercek klasore cevirir.
' ============================================================================

Private Const ADO_TIP_METIN As Long = 2
Private Const ADO_TIP_IKILI As Long = 1
Private Const ADO_YAZ_YOKET As Long = 2      ' adSaveCreateOverWrite
Private Const ADO_YAZ_YOKSA As Long = 1      ' adSaveCreateNotExist


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

Public Function DosyaAdiUzantisiz(ByVal yol As String) As String
    Dim ad As String, p As Long
    ad = DosyaAdi(yol)
    p = InStrRev(ad, ".")
    If p <= 1 Then DosyaAdiUzantisiz = ad Else DosyaAdiUzantisiz = Left$(ad, p - 1)
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
'  YerelYol -- OneDrive/SharePoint adresini diskteki gercek klasore cevirir.
'
'  Ornek:  https://kurum-my.sharepoint.com/personal/x/Documents/Masaüstü/kaizen
'          -> C:\Users\x\OneDrive - Kurum\Masaüstü\kaizen
'
'  Yontem: adresin sonundan baslayarak gitgide uzayan parcalar, bilinen yerel
'  OneDrive koklerinin altinda aranir; diskte gercekten var olan ilk eslesme
'  dondurulur. Eslesme bulunamazsa adres oldugu gibi geri verilir -- cagiran
'  taraf o zaman anlamli bir hata gosterir.
'
'  UNC yollari (\\sunucu\paylasim\...) dokunulmadan gecer; uretim ortami budur.
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

    For Each ad In Array("OneDrive", "OneDriveCommercial", "OneDriveConsumer")
        deger = Environ$(CStr(ad))
        If Len(deger) > 0 Then KokEkle c, deger
    Next ad

    ' "OneDrive - <Kurum>" bicimindeki klasorler ortam degiskeninde her zaman
    ' gorunmez; kullanici profilinin altini da tararız.
    profil = Environ$("USERPROFILE")
    If Len(profil) > 0 Then
        klasorAdi = Dir$(profil & "\OneDrive*", vbDirectory)
        Do While Len(klasorAdi) > 0
            If klasorAdi <> "." And klasorAdi <> ".." Then
                KokEkle c, profil & "\" & klasorAdi
            End If
            klasorAdi = Dir$
        Loop
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
'  UTF-8 OKUMA / YAZMA
' ###########################################################################

' UTF-8 (BOM'suz) yazar. Once metin akisina yazilir, sonra ilk uc bayt (BOM)
' atlanarak ikili akisa kopyalanir; ADODB.Stream'in BOM'u kaldirma yolu budur.
'
' uzerineYaz:=False verilirse dosya YALNIZCA yoksa olusturulur; varsa islem
' hata verir ve mevcut dosyaya dokunulmaz.
Public Sub UTF8Yaz(ByVal dosyaYolu As String, ByVal metin As String, _
                   Optional ByVal uzerineYaz As Boolean = True)
    Dim metinAkis As Object, ikiliAkis As Object
    Dim kayitKipi As Long

    If uzerineYaz Then kayitKipi = ADO_YAZ_YOKET Else kayitKipi = ADO_YAZ_YOKSA

    Set metinAkis = CreateObject("ADODB.Stream")
    metinAkis.Type = ADO_TIP_METIN
    metinAkis.Charset = "utf-8"
    metinAkis.Open
    metinAkis.WriteText metin

    metinAkis.Position = 0
    metinAkis.Type = ADO_TIP_IKILI
    metinAkis.Position = 3                      ' UTF-8 BOM'unu atla

    Set ikiliAkis = CreateObject("ADODB.Stream")
    ikiliAkis.Type = ADO_TIP_IKILI
    ikiliAkis.Open
    metinAkis.CopyTo ikiliAkis
    ikiliAkis.SaveToFile dosyaYolu, kayitKipi
    ikiliAkis.Close
    metinAkis.Close

    Set ikiliAkis = Nothing
    Set metinAkis = Nothing
End Sub

Public Function UTF8Oku(ByVal dosyaYolu As String) As String
    Dim akis As Object
    Set akis = CreateObject("ADODB.Stream")
    akis.Type = ADO_TIP_METIN
    akis.Charset = "utf-8"
    akis.Open
    akis.LoadFromFile dosyaYolu
    UTF8Oku = akis.ReadText(-1)                 ' adReadAll
    akis.Close
    Set akis = Nothing
End Function


' ---------------------------------------------------------------------------
'  TAM KAYIT YAZMA -- iki guvence, tek adimda
'
'  1) UZERINE YAZMAZ. Dosya "varsa oluşturma" kipiyle acilir; ayni adda bir
'     dosya varsa islem basarisiz olur ve mevcut dosyaya DOKUNULMAZ. Kisa
'     oneri numaralariyla bu sarttir: iki kullanici ayni numarayi uretirse
'     ikincisi digerinin onerisini yok edemez, yeni numarayla yeniden dener.
'
'  2) YARIM DOSYA OKUNMAZ. Kaydin sonuna bir "kayit_sonu" satiri eklenir.
'     Yazma yarida kesilirse (ag koptu, Excel kapandi) bu satir olusmaz ve
'     okuyucu dosyayi yok sayar.
'
'  NEDEN ".tmp + yeniden adlandirma" DEGIL:
'  Onceki tasarim once ".tmp" yazip sonra yeniden adlandiriyordu. Yeniden
'  adlandirma, kaynak dosya uzerinde SILME yetkisi ister. Oysa
'  "yonetim\oneriler\" bir birakma kutusudur: personel oraya yazabilir ama
'  silemez ve icerigini goremez (bkz. KURULUM.md). O izinlerle yeniden adlandirma
'  isletim sistemi tarafindan reddediliyor. Dosyayi dogrudan son adiyla
'  olusturmak ayni iki guvenceyi yalnizca YAZMA yetkisiyle saglar.
' ---------------------------------------------------------------------------
Public Sub TamKayitYaz(ByVal hedefDosya As String, ByVal metin As String)
    Dim aciklama As String
    If Not TamKayitYazmayiDene(hedefDosya, metin, aciklama) Then
        Err.Raise vbObjectError + 910, "modDosyaIO.TamKayitYaz", _
                  "Dosya yazılamadı: " & hedefDosya & vbCrLf & aciklama
    End If
End Sub

' Yazabildiyse True doner. False donmesinin iki nedeni olabilir: ayni adda
' dosya vardir ya da klasore yazma izni yoktur. ADODB ikisi icin de ayni
' hatayi verdiginden ayirt edilemez; cagiran taraf birkac kez yeniden dener,
' hepsi basarisiz olursa izin sorunu oldugunu bildirir (bkz. modGonderim).
Public Function TamKayitYazmayiDene(ByVal hedefDosya As String, _
                                    ByVal metin As String, _
                                    Optional ByRef hataAciklamasi As String) As Boolean
    On Error GoTo Hata

    KlasorZinciriOlustur UstKlasor(hedefDosya)
    UTF8Yaz hedefDosya, metin & modAyar.ALAN_SON & "=1" & vbLf, uzerineYaz:=False

    hataAciklamasi = ""
    TamKayitYazmayiDene = True
    Exit Function

Hata:
    hataAciklamasi = Err.Description
    TamKayitYazmayiDene = False
End Function


' ###########################################################################
'  KAYIT BICIMI  (anahtar=deger, UTF-8)
' ###########################################################################

' Cok satirli metni tek satira sikistirir. Satir sonlari sabit bir belirtece
' cevrilir; boylece her alan dosyada tam olarak bir satir kaplar.
Public Function DegerKacisla(ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    DegerKacisla = Replace(t, vbLf, modAyar.SATIR_BELIRTEC)
End Function

Public Function DegerCoz(ByVal s As String) As String
    DegerCoz = Replace(s, modAyar.SATIR_BELIRTEC, vbLf)
End Function

' Sirali bir anahtar dizisi ve sozluk alarak kayit metnini uretir.
Public Function KayitMetniUret(ByVal sozluk As Object, ByVal anahtarlar As Variant) As String
    Dim i As Long, satirlar As String
    For i = LBound(anahtarlar) To UBound(anahtarlar)
        satirlar = satirlar & CStr(anahtarlar(i)) & "=" & _
                   DegerKacisla(CStr(sozluk(CStr(anahtarlar(i))))) & vbLf
    Next i
    KayitMetniUret = satirlar
End Function

' Kayit dosyasini okuyup Scripting.Dictionary dondurur.
' Bilinmeyen ya da bozuk satirlar sessizce atlanir: elle onarilmis bir dosya
' yuzunden butun konsolidasyon durmamalidir.
Public Function KayitOku(ByVal dosyaYolu As String) As Object
    Dim sozluk As Object, ham As String
    Dim satirlar() As String, i As Long, p As Long
    Dim anahtar As String, deger As String

    Set sozluk = CreateObject("Scripting.Dictionary")
    sozluk.CompareMode = 1                      ' vbTextCompare

    On Error GoTo Hata
    ham = UTF8Oku(dosyaYolu)
    On Error GoTo 0

    ham = Replace(ham, vbCrLf, vbLf)
    ham = Replace(ham, vbCr, vbLf)
    satirlar = Split(ham, vbLf)

    For i = LBound(satirlar) To UBound(satirlar)
        p = InStr(satirlar(i), "=")
        If p > 1 Then
            anahtar = Trim$(Left$(satirlar(i), p - 1))
            deger = Mid$(satirlar(i), p + 1)
            If Len(anahtar) > 0 Then sozluk(anahtar) = DegerCoz(deger)
        End If
    Next i

    Set KayitOku = sozluk
    Exit Function

Hata:
    Set KayitOku = sozluk                       ' okunamayan dosya = bos kayit
End Function

' ---------------------------------------------------------------------------
'  ZamanDamgasi -- "yyyymmddhhnnss" + iki haneli salise (10 ms cozunurluk)
'
'  Olay dosyalarinin sirasi ADLARINDAN okunur; sira = zaman sirasi kabul edilir.
'  Saniye cozunurlugu yetmez: ayni saniye icinde yazilan iki olayin sirasini
'  addaki rastgele ek belirlerdi ve durum yanlis turetilebilirdi (once
'  "Planlandı" sonra "Pilot Uygulamada" yazilmisken tersi okunabilirdi).
'  Salise eklenmesi bu belirsizligi pratikte ortadan kaldirir.
' ---------------------------------------------------------------------------
Public Function ZamanDamgasi() As String
    Dim t As Double, salise As Long
    t = Timer                                  ' gece yarisindan beri gecen saniye
    salise = Int((t - Int(t)) * 100)
    If salise < 0 Then salise = 0
    If salise > 99 Then salise = 99
    ZamanDamgasi = Format$(Now, "yyyymmddhhnnss") & Format$(salise, "00")
End Function

' Dosya adlarindaki rastgele ek. Zaman damgasi tek basina yetmez: ayni salise
' icinde iki kayit olusabilir. Hem gonderim hem degerlendirme tarafi kullanir.
Public Function RastgeleOnaltilik(ByVal hane As Long) As String
    Static tohumAtildi As Boolean
    Dim s As String, i As Long

    If Not tohumAtildi Then
        Randomize
        tohumAtildi = True
    End If

    For i = 1 To hane
        s = s & Mid$("0123456789ABCDEF", Int(Rnd() * 16) + 1, 1)
    Next i
    RastgeleOnaltilik = s
End Function

' ---------------------------------------------------------------------------
'  KisaRastgele -- oneri numarasinin okunabilir rastgele eki.
'
'  Alfabeden 0/O ve 1/I cikarilmistir: numara telefonda soylenirken ya da
'  elle yazilirken karistirilmasin. Geriye 32 karakter kalir; uc hane
'  32.768 farkli deger demektir. Ayni gun icinde catisma pratikte olmaz,
'  olsa bile cagiran taraf yeni numara uretip yeniden dener.
' ---------------------------------------------------------------------------
Public Function KisaRastgele(ByVal hane As Long) As String
    Const ALFABE As String = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ"
    Static tohumAtildi As Boolean
    Dim s As String, i As Long

    If Not tohumAtildi Then
        Randomize
        tohumAtildi = True
    End If

    For i = 1 To hane
        s = s & Mid$(ALFABE, Int(Rnd() * Len(ALFABE)) + 1, 1)
    Next i
    KisaRastgele = s
End Function

' Kayit eksiksiz yazilmis mi? Son satir ("kayit_sonu") yoksa dosya yarim
' kalmistir; okuyucu boyle bir kaydi yok saymalidir.
Public Function KayitTamMi(ByVal sozluk As Object) As Boolean
    If sozluk Is Nothing Then Exit Function
    KayitTamMi = sozluk.Exists(modAyar.ALAN_SON)
End Function

' Sozlukten guvenli okuma (anahtar yoksa varsayilan doner).
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


' ###########################################################################
'  KLASOR TARAMA
' ###########################################################################

' Bir klasoru ve yil alt klasorlerini tarar, verilen uzantiya sahip dosyalarin
' tam yollarini ADA GORE SIRALI dondurur.
' Dosya adlari zaman damgasiyla basladigi icin ad sirasi = zaman sirasi.
' ".tmp" dosyalari (yarim yazilmis kayitlar) dogal olarak disarida kalir.
Public Function DosyalariTara(ByVal kokKlasor As String, _
                              Optional ByVal uzanti As String = ".txt") As Collection
    Dim sonuc As New Collection
    Dim altKlasorler As New Collection
    Dim ad As String, alt As Variant
    Dim liste() As String, sayi As Long, i As Long

    Set DosyalariTara = sonuc
    If Not KlasorVarMi(kokKlasor) Then Exit Function

    ' Once alt klasorleri topla (Dir$ ic ice cagrilamaz).
    ad = Dir$(YolTemizle(kokKlasor) & "\*", vbDirectory)
    Do While Len(ad) > 0
        If ad <> "." And ad <> ".." Then
            If (GetAttr(YolTemizle(kokKlasor) & "\" & ad) And vbDirectory) = vbDirectory Then
                altKlasorler.Add YolTemizle(kokKlasor) & "\" & ad
            End If
        End If
        ad = Dir$
    Loop

    ' Bu klasordeki dosyalar
    ReDim liste(0 To 0)
    sayi = 0
    ad = Dir$(YolTemizle(kokKlasor) & "\*" & uzanti, vbNormal)
    Do While Len(ad) > 0
        If StrComp(Right$(ad, Len(uzanti)), uzanti, vbTextCompare) = 0 Then
            ReDim Preserve liste(0 To sayi)
            liste(sayi) = YolTemizle(kokKlasor) & "\" & ad
            sayi = sayi + 1
        End If
        ad = Dir$
    Loop

    If sayi > 0 Then
        DiziSirala liste, sayi
        For i = 0 To sayi - 1
            sonuc.Add liste(i)
        Next i
    End If

    ' Alt klasorler (yil klasorleri) -- ozyinelemeli
    For Each alt In altKlasorler
        Dim altSonuc As Collection, v As Variant
        Set altSonuc = DosyalariTara(CStr(alt), uzanti)
        For Each v In altSonuc
            sonuc.Add v
        Next v
    Next alt

    Set DosyalariTara = sonuc
End Function

' Kucuk dosya listeleri icin ekleme siralamasi yeterlidir.
Private Sub DiziSirala(ByRef liste() As String, ByVal sayi As Long)
    Dim i As Long, j As Long, gecici As String
    For i = 1 To sayi - 1
        gecici = liste(i)
        j = i - 1
        Do While j >= 0
            If StrComp(liste(j), gecici, vbTextCompare) > 0 Then
                liste(j + 1) = liste(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        liste(j + 1) = gecici
    Next i
End Sub

' Bir Collection'i ada gore sirali String dizisine cevirir.
Public Function KoleksiyonuSirala(ByVal c As Collection) As Variant
    Dim liste() As String, i As Long, v As Variant
    If c Is Nothing Then
        KoleksiyonuSirala = Array()
        Exit Function
    End If
    If c.Count = 0 Then
        KoleksiyonuSirala = Array()
        Exit Function
    End If
    ReDim liste(0 To c.Count - 1)
    i = 0
    For Each v In c
        liste(i) = CStr(v)
        i = i + 1
    Next v
    DiziSirala liste, c.Count
    KoleksiyonuSirala = liste
End Function
