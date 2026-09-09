Attribute VB_Name = "modKonsolide"
Option Explicit

' ============================================================================
'  modKonsolide -- depo sayfalarindan ekrandaki tabloya
'
'  EKRANDA GORUNEN HICBIR SEY ELLE GIRILMEZ. Liste de pano da, kitabin
'  icindeki iki depo sayfasindan (Oneriler, Olaylar) her yenilemede bastan
'  turetilir. Gizli "Veri" sayfasi bu turetmenin sonucudur; bir onbellektir,
'  kaynak degildir.
'
'  Guncel durum nasil bulunur:
'    1) "Oneriler" satirlari okunur -- bunlar DEGISMEZDIR, guncellenmez.
'    2) "Olaylar" satirlari SATIR SIRASIYLA (= zaman sirasiyla) uzerlerine
'       oynatilir. Her olay yalnizca DOLDURDUGU alanlari gunceller; boylece
'       sadece durum degistiren bir olay, daha once yazilmis karar notunu
'       silmez.
'
'  ---------------------------------------------------------------------
'  YEREL KOPYA ve "Yenile"
'
'  Ekip kitabi salt okunur acilir ve gun boyu acik kalir; bu sirada personel
'  diskteki dosyaya yeni satirlar ekler. Bellekteki kopya bunlari kendiliginden
'  gormez. "Önerileri Yenile" once modDepo.DepoyuCek ile diskteki depo
'  sayfalarini bu kitaba kopyalar, sonra tabloyu bastan kurar.
'
'  Oturum acilisinda cekmeye gerek yoktur: kitap saniyeler once diskten
'  yuklenmistir.
' ============================================================================

' --- Gizli "Veri" sayfasinin sutun duzeni --------------------------------
Public Const V_ONERI_NO As Long = 1
Public Const V_TARIH As Long = 2
Public Const V_AD_SOYAD As Long = 3
Public Const V_SICIL_NO As Long = 4
Public Const V_MEVCUT As Long = 5
Public Const V_BASLIK As Long = 6
Public Const V_COZUM As Long = 7
Public Const V_FAYDA As Long = 8
Public Const V_DURUM As Long = 9
Public Const V_ILK_OLAY As Long = 10
Public Const V_SON_OLAY As Long = 11
Public Const V_DEGERLENDIREN As Long = 12
Public Const V_KARAR_NOTU As Long = 13
Public Const V_OLAY_SAYISI As Long = 14
Public Const V_GONDEREN_KULLANICI As Long = 15
Public Const V_KAYNAK_SATIR As Long = 16
Public Const V_SUTUN_SAYISI As Long = 16

' --- Liste tablosu -------------------------------------------------------
Public Const LISTE_ILK_SATIR As Long = 9
Public Const LISTE_ILK_SUTUN As Long = 2      ' B
Public Const LISTE_SON_SUTUN As Long = 6      ' F
Public Const LISTE_AZAMI_SATIR As Long = 2000


' ###########################################################################
'  DUGME EYLEMLERI
' ###########################################################################

' "Sisteme Gir" dugmesi -- Giris sayfasi
'
' Oturum PANO ile acilir: ekip once genel resmi gorur, ayrintiya listeden
' iner. Liste bir dugme uzaktadir.
Public Sub SistemeGir()
    Dim hataMetni As String

    If Not modUI.SifreDogrula(modAyar.SIFRE_YONETIM, "Proje Öneri Yönetimi") Then Exit Sub
    modUI.OturumAc Array(modUI.SAYFA_PANO, modUI.SAYFA_LISTE, _
                         modUI.SAYFA_DEGERLENDIRME), modUI.SAYFA_PANO

    On Error GoTo Hata
    modUI.HizliModAc
    OzetYaz YerelYenile()
    modUI.HizliModKapa
    Exit Sub

Hata:
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    modUI.HizliModKapa
    modUI.Hata "Ekranlar hazırlanamadı." & vbCrLf & vbCrLf & hataMetni, _
               "Proje Öneri Yönetimi"
End Sub

' "Çıkış" dugmesi
Public Sub Cikis()
    modUI.OturumKapat
End Sub

' "Önerileri Yenile" dugmesi -- diskteki depoyu ceker, sonra tabloyu kurar.
Public Sub OnerileriYenile()
    Dim adet As Long
    Dim hataMetni As String

    On Error GoTo Hata
    modUI.HizliModAc
    DurumCubugu "Öneriler okunuyor..."

    modDepo.DepoyuCek ThisWorkbook
    adet = YerelYenile()

    DurumCubugu ""
    modUI.HizliModKapa
    OzetYaz adet
    Exit Sub

Hata:
    ' Aciklama ILK is olarak alinir: sonraki her cagri Err'i temizler.
    hataMetni = "Hata " & Err.Number & ": " & Err.Description
    DurumCubugu ""
    modUI.HizliModKapa
    modUI.Hata "Öneriler okunamadı." & vbCrLf & vbCrLf & hataMetni & _
               vbCrLf & vbCrLf & _
               "Yönetim kitabına o anda başka bir kullanıcı yazıyor " & _
               "olabilir; birkaç saniye sonra yeniden deneyin.", _
               "Yenileme hatası"
End Sub

' Diski OKUMADAN, kitaptaki depo sayfalarindan ekrani kurar.
Public Function YerelYenile() As Long
    YerelYenile = VeriyiKur()
    ListeyiCiz
    modPano.PanoyuYenile
End Function

' Durum cubugu yalnizca gorsel bir bilgidir; bir hata isleyicisi icinden de
' cagrildigi icin tamamen hataya dayanikli olmalidir.
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

' Listenin ustundeki tek satirlik ozet: kac oneri var, kaci bekliyor, ne zaman
' yenilendi. Public'tir; testler tek basina calistirabilsin diye.
Public Sub OzetYaz(ByVal adet As Long)
    Dim ws As Object, bekleyen As Long, i As Long, veri As Variant

    Set ws = ListeSayfasi()
    veri = VeriDizisi()
    If IsEmpty(veri) Then
        bekleyen = 0
    Else
        For i = LBound(veri, 1) To UBound(veri, 1)
            If modModel.BekliyorMu(CStr(veri(i, V_DURUM))) Then bekleyen = bekleyen + 1
        Next i
    End If

    modUI.KorumaKapa ws
    ws.Range("liste_ozet").Value = _
        adet & " öneri okundu   ·   " & bekleyen & " tanesi değerlendirme bekliyor" & _
        "   ·   son yenileme " & Format$(Now, "dd.mm.yyyy hh:nn")
    modUI.KorumaAc ws
End Sub


' ###########################################################################
'  1. DEPO SAYFALARINDAN "Veri" SAYFASINA
' ###########################################################################

Public Function VeriyiKur() As Long
    Dim ws As Object
    Dim oneriler As Variant, olaylar As Variant
    Dim indeks As Object
    Dim veri() As Variant
    Dim n As Long, i As Long, r As Long, no As String

    Set ws = VeriSayfasi()
    oneriler = modDepo.TabloOku(OnerilerSayfasi(), modDepo.O_SUTUN_SAYISI)

    modUI.KorumaKapa ws
    ws.Range(ws.Cells(2, 1), _
             ws.Cells(LISTE_AZAMI_SATIR + 200, V_SUTUN_SAYISI)).ClearContents

    If IsEmpty(oneriler) Then
        ws.Range("veri_adet").Value = 0
        modUI.KorumaAc ws
        VeriyiKur = 0
        Exit Function
    End If

    n = UBound(oneriler, 1)
    ReDim veri(1 To n, 1 To V_SUTUN_SAYISI)
    Set indeks = CreateObject("Scripting.Dictionary")
    indeks.CompareMode = 1

    ' --- Gonderimler ------------------------------------------------------
    i = 0
    For r = 1 To n
        no = Trim$(CStr(oneriler(r, modDepo.O_ONERI_NO) & ""))

        ' Numarasiz satir yok sayilir: elle bozulmus ya da yarim birakilmis
        ' bir satir yuzunden butun konsolidasyon durmamalidir.
        If Len(no) > 0 Then
            If Not indeks.Exists(no) Then
                i = i + 1
                indeks(no) = i
                veri(i, V_ONERI_NO) = no
                veri(i, V_TARIH) = oneriler(r, modDepo.O_TARIH)
                veri(i, V_AD_SOYAD) = oneriler(r, modDepo.O_AD_SOYAD)
                veri(i, V_SICIL_NO) = oneriler(r, modDepo.O_SICIL_NO)
                veri(i, V_MEVCUT) = oneriler(r, modDepo.O_MEVCUT_DURUM)
                veri(i, V_BASLIK) = oneriler(r, modDepo.O_ONERI_BASLIGI)
                veri(i, V_COZUM) = oneriler(r, modDepo.O_COZUM_ONERISI)
                veri(i, V_FAYDA) = oneriler(r, modDepo.O_BEKLENEN_FAYDA)
                veri(i, V_DURUM) = modModel.DURUM_YENI   ' durum yalnizca olaylardan gelir
                veri(i, V_ILK_OLAY) = ""
                veri(i, V_SON_OLAY) = ""
                veri(i, V_DEGERLENDIREN) = ""
                veri(i, V_KARAR_NOTU) = ""
                veri(i, V_OLAY_SAYISI) = 0
                veri(i, V_GONDEREN_KULLANICI) = oneriler(r, modDepo.O_GONDEREN_KULLANICI)
                veri(i, V_KAYNAK_SATIR) = r + 1          ' depo sayfasindaki satir
            End If
        End If
    Next r
    n = i

    ' --- Olaylar (ekle-only gecmis) --------------------------------------
    olaylar = modDepo.TabloOku(OlaylarSayfasi(), modDepo.E_SUTUN_SAYISI)
    If Not IsEmpty(olaylar) Then
        For r = LBound(olaylar, 1) To UBound(olaylar, 1)
            no = Trim$(CStr(olaylar(r, modDepo.E_ONERI_NO) & ""))
            If Len(no) > 0 Then
                If indeks.Exists(no) Then
                    OlayiUygula veri, CLng(indeks(no)), olaylar, r
                End If
            End If
        Next r
    End If

    If n > 0 Then
        ws.Cells(2, 1).Resize(n, V_SUTUN_SAYISI).Value = veri
    End If
    ws.Range("veri_adet").Value = n
    modUI.KorumaAc ws

    VeriyiKur = n
End Function


' Bir olay satirini kayda uygular. Yalnizca DOLU alanlar yazilir: kismi bir
' olay (ornegin yalnizca durum degisikligi) onceki karar notunu silmemelidir.
Private Sub OlayiUygula(ByRef veri() As Variant, ByVal i As Long, _
                        ByVal olaylar As Variant, ByVal r As Long)
    Dim s As String

    s = Trim$(CStr(olaylar(r, modDepo.E_YENI_DURUM) & ""))
    If Len(s) > 0 Then
        If modModel.DurumGecerliMi(s) Then veri(i, V_DURUM) = s
    End If

    s = CStr(olaylar(r, modDepo.E_KARAR_NOTU) & "")
    If Len(s) > 0 Then veri(i, V_KARAR_NOTU) = s

    s = Trim$(CStr(olaylar(r, modDepo.E_OLAY_TARIHI) & ""))
    If Len(s) > 0 Then
        If Len(CStr(veri(i, V_ILK_OLAY))) = 0 Then veri(i, V_ILK_OLAY) = s
        veri(i, V_SON_OLAY) = s
    End If

    s = Trim$(CStr(olaylar(r, modDepo.E_DEGERLENDIREN_KULLANICI) & ""))
    If Len(s) > 0 Then veri(i, V_DEGERLENDIREN) = s

    veri(i, V_OLAY_SAYISI) = CLng(veri(i, V_OLAY_SAYISI)) + 1
End Sub


' ###########################################################################
'  2. "Veri" SAYFASINDAN LISTEYE
' ###########################################################################

Public Sub ListeyiCiz()
    Dim ws As Object, veri As Variant
    Dim sira() As Long
    Dim n As Long, i As Long, r As Long
    Dim tablo() As Variant

    Set ws = ListeSayfasi()
    modUI.KorumaKapa ws

    ' Onceki icerigi temizle (bicimler sayfada kalir, yalnizca deger silinir).
    ws.Range(ws.Cells(LISTE_ILK_SATIR, LISTE_ILK_SUTUN), _
             ws.Cells(LISTE_AZAMI_SATIR, LISTE_SON_SUTUN)).ClearContents
    FiltreyiKaldir ws

    veri = VeriDizisi()
    If IsEmpty(veri) Then
        modUI.KorumaAc ws
        Exit Sub
    End If

    n = UBound(veri, 1)
    sira = SiralamaDizisi(veri, n)

    ReDim tablo(1 To n, 1 To LISTE_SON_SUTUN - LISTE_ILK_SUTUN + 1)
    For r = 1 To n
        i = sira(r)
        tablo(r, 1) = veri(i, V_ONERI_NO)
        tablo(r, 2) = TarihGoster(CStr(veri(i, V_TARIH)))
        tablo(r, 3) = veri(i, V_AD_SOYAD)
        tablo(r, 4) = veri(i, V_BASLIK)
        tablo(r, 5) = veri(i, V_DURUM)
    Next r

    ws.Cells(LISTE_ILK_SATIR, LISTE_ILK_SUTUN).Resize(n, UBound(tablo, 2)).Value = tablo
    FiltreKur ws, n
    modUI.KorumaAc ws
End Sub

Private Function TarihGoster(ByVal isoTarih As String) As String
    ' "2026-09-04T10:11:51" -> "04.09.2026"
    If Len(isoTarih) >= 10 Then
        TarihGoster = Mid$(isoTarih, 9, 2) & "." & Mid$(isoTarih, 6, 2) & "." & _
                      Left$(isoTarih, 4)
    Else
        TarihGoster = isoTarih
    End If
End Function

' Varsayilan sira: once bekleyenler (PDCA akisindaki sirayla), her grup
' kendi icinde eskiden yeniye. Değerlendirme ekibi ekrani actiginda sirada ne
' oldugunu ustte gorur.
Private Function SiralamaDizisi(ByVal veri As Variant, ByVal n As Long) As Long()
    Dim sira() As Long, anahtar() As String
    Dim i As Long, j As Long, gecici As Long, gAnahtar As String

    ReDim sira(1 To n)
    ReDim anahtar(1 To n)
    For i = 1 To n
        sira(i) = i
        anahtar(i) = Format$(modModel.DurumSirasi(CStr(veri(i, V_DURUM))), "00") & _
                     "|" & CStr(veri(i, V_TARIH)) & "|" & CStr(veri(i, V_ONERI_NO))
    Next i

    For i = 2 To n
        gecici = sira(i)
        gAnahtar = anahtar(gecici)
        j = i - 1
        Do While j >= 1
            If StrComp(anahtar(sira(j)), gAnahtar, vbTextCompare) > 0 Then
                sira(j + 1) = sira(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        sira(j + 1) = gecici
    Next i

    SiralamaDizisi = sira
End Function

Private Sub FiltreyiKaldir(ByVal ws As Object)
    On Error Resume Next
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    On Error GoTo 0
End Sub

Private Sub FiltreKur(ByVal ws As Object, ByVal n As Long)
    On Error Resume Next
    If n > 0 Then
        ws.Range(ws.Cells(LISTE_ILK_SATIR - 1, LISTE_ILK_SUTUN), _
                 ws.Cells(LISTE_ILK_SATIR - 1 + n, LISTE_SON_SUTUN)).AutoFilter
    End If
    On Error GoTo 0
End Sub


' ###########################################################################
'  ERISIM YARDIMCILARI
' ###########################################################################

Public Function VeriSayfasi() As Object
    Set VeriSayfasi = ThisWorkbook.Worksheets(modUI.SAYFA_VERI)
End Function

Public Function ListeSayfasi() As Object
    Set ListeSayfasi = ThisWorkbook.Worksheets(modUI.SAYFA_LISTE)
End Function

' Bu kitaptaki depo sayfalari. Diskteki dosyanin ayni adli sayfalarinin yerel
' kopyasidir; modDepo.DepoyuCek her yenilemede uzerine yazar.
Public Function OnerilerSayfasi() As Object
    Set OnerilerSayfasi = ThisWorkbook.Worksheets(modDepo.SAYFA_ONERILER)
End Function

Public Function OlaylarSayfasi() As Object
    Set OlaylarSayfasi = ThisWorkbook.Worksheets(modDepo.SAYFA_OLAYLAR)
End Function

' Listede uzerine tiklanan satirin oneri numarasi. Tablo disinda bir yer
' seciliyse bos doner.
Public Function SeciliOneriNo() As String
    Dim ws As Object, satir As Long, deger As String

    Set ws = ListeSayfasi()
    satir = ActiveCell.Row
    If satir < LISTE_ILK_SATIR Or satir > LISTE_AZAMI_SATIR Then Exit Function

    deger = Trim$(CStr(ws.Cells(satir, LISTE_ILK_SUTUN).Value & ""))
    If Len(deger) = 0 Then Exit Function
    SeciliOneriNo = deger
End Function

Public Function VeriAdedi() As Long
    Dim v As Variant
    On Error Resume Next
    v = VeriSayfasi().Range("veri_adet").Value
    On Error GoTo 0
    If IsNumeric(v) Then VeriAdedi = CLng(v) Else VeriAdedi = 0
End Function

' Veri sayfasinin tamamini bir dizi olarak dondurur. Kayit yoksa Empty doner.
Public Function VeriDizisi() As Variant
    Dim n As Long, ws As Object

    n = VeriAdedi()
    If n <= 0 Then
        VeriDizisi = Empty
        Exit Function
    End If

    Set ws = VeriSayfasi()
    If n = 1 Then
        ' Tek satirda Range.Value dizi dondurmez; elle iki boyutlu yapariz.
        Dim tek() As Variant, c As Long
        ReDim tek(1 To 1, 1 To V_SUTUN_SAYISI)
        For c = 1 To V_SUTUN_SAYISI
            tek(1, c) = ws.Cells(2, c).Value
        Next c
        VeriDizisi = tek
    Else
        VeriDizisi = ws.Cells(2, 1).Resize(n, V_SUTUN_SAYISI).Value
    End If
End Function

' Oneri numarasina gore Veri dizisindeki satiri bulur (0 = bulunamadi).
Public Function SatirBul(ByVal veri As Variant, ByVal oneriNo As String) As Long
    Dim i As Long
    If IsEmpty(veri) Then Exit Function
    For i = LBound(veri, 1) To UBound(veri, 1)
        If StrComp(CStr(veri(i, V_ONERI_NO)), oneriNo, vbTextCompare) = 0 Then
            SatirBul = i
            Exit Function
        End If
    Next i
End Function
