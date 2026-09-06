Attribute VB_Name = "modTasarim"
Option Explicit

' ============================================================================
'  modTasarim -- kurumsal tasarim sabitleri
'
'  Buradaki hex degerleri kaynak\tasarim.py ile BIREBIR ayni olmak zorundadir.
'  kur.py uretimden once satir sonundaki '#HEX etiketlerini okuyup karsilastirir;
'  uyusmazlik varsa uretim durur. Etiketleri silmeyin.
' ============================================================================

' --- Kurumsal palet -------------------------------------------------------
Public Const CLR_ANA_LACIVERT As String = "002D62"    '#HEX ANA_LACIVERT
Public Const CLR_DERIN_LACIVERT As String = "001F44"  '#HEX DERIN_LACIVERT
Public Const CLR_KURUMSAL_MAVI As String = "1B5FAA"   '#HEX KURUMSAL_MAVI
Public Const CLR_ACIK_MAVI As String = "D6E4F0"       '#HEX ACIK_MAVI
Public Const CLR_SOLUK_MAVI As String = "9FB6D4"      '#HEX SOLUK_MAVI
Public Const CLR_BUZ_ZEMIN As String = "F4F7FB"       '#HEX BUZ_ZEMIN
Public Const CLR_BEYAZ As String = "FFFFFF"           '#HEX BEYAZ
Public Const CLR_CIZGI_GRI As String = "D9DEE6"       '#HEX CIZGI_GRI
Public Const CLR_GOLGE_GRI As String = "C6CDD6"       '#HEX GOLGE_GRI
Public Const CLR_ALAN_CIZGI As String = "B7C4D6"      '#HEX ALAN_CIZGI
Public Const CLR_METIN_KOYU As String = "1A2433"      '#HEX METIN_KOYU
Public Const CLR_METIN_GRI As String = "5A6572"       '#HEX METIN_GRI
Public Const CLR_ONAY_ZEMIN As String = "D2EAD9"      '#HEX ONAY_ZEMIN
Public Const CLR_ONAY_YAZI As String = "1E7B4D"       '#HEX ONAY_YAZI
Public Const CLR_UYARI_ZEMIN As String = "F6DADA"     '#HEX UYARI_ZEMIN
Public Const CLR_UYARI_YAZI As String = "9A2E2E"      '#HEX UYARI_YAZI

' --- Tipografi ------------------------------------------------------------
Public Const FONT_AILE As String = "Segoe UI"
Public Const FONT_BASLIK_AILE As String = "Segoe UI Semibold"


' ---------------------------------------------------------------------------
'  HexRGB -- "002D62" gibi bir hex kodu Excel'in Long renk degerine cevirir.
'  Excel renkleri BGR sirasindadir; bu yuzden mavi en yuksek bayta gider.
' ---------------------------------------------------------------------------
Public Function HexRGB(ByVal hexKod As String) As Long
    Dim k As String
    k = Replace(hexKod, "#", "")
    If Len(k) <> 6 Then
        HexRGB = vbBlack
        Exit Function
    End If
    HexRGB = RGB(CLng("&H" & Mid$(k, 1, 2)), _
                 CLng("&H" & Mid$(k, 3, 2)), _
                 CLng("&H" & Mid$(k, 5, 2)))
End Function


' ---------------------------------------------------------------------------
'  Durum renkleri
' ---------------------------------------------------------------------------
Public Function DurumZeminHex(ByVal durum As String) As String
    Select Case durum
        Case "Yeni":               DurumZeminHex = "D6E4F0"   '#HEX DURUM_Yeni_ZEMIN
        Case "Değerlendirmede":    DurumZeminHex = "FCEBCF"   '#HEX DURUM_Değerlendirmede_ZEMIN
        Case "Planlandı":          DurumZeminHex = "E4DBF5"   '#HEX DURUM_Planlandı_ZEMIN
        Case "Pilot Uygulamada":   DurumZeminHex = "D5EEF0"   '#HEX DURUM_Pilot Uygulamada_ZEMIN
        Case "Ölçümleniyor":       DurumZeminHex = "DDEBD9"   '#HEX DURUM_Ölçümleniyor_ZEMIN
        Case "Standartlaştırıldı": DurumZeminHex = "D2EAD9"   '#HEX DURUM_Standartlaştırıldı_ZEMIN
        Case "Beklemede":          DurumZeminHex = "EDEFF2"   '#HEX DURUM_Beklemede_ZEMIN
        Case "Reddedildi":         DurumZeminHex = "F6DADA"   '#HEX DURUM_Reddedildi_ZEMIN
        Case Else:                 DurumZeminHex = CLR_BEYAZ
    End Select
End Function

Public Function DurumYaziHex(ByVal durum As String) As String
    Select Case durum
        Case "Yeni":               DurumYaziHex = "002D62"    '#HEX DURUM_Yeni_YAZI
        Case "Değerlendirmede":    DurumYaziHex = "8A5A00"    '#HEX DURUM_Değerlendirmede_YAZI
        Case "Planlandı":          DurumYaziHex = "4B3A8C"    '#HEX DURUM_Planlandı_YAZI
        Case "Pilot Uygulamada":   DurumYaziHex = "0E6470"    '#HEX DURUM_Pilot Uygulamada_YAZI
        Case "Ölçümleniyor":       DurumYaziHex = "3A6B35"    '#HEX DURUM_Ölçümleniyor_YAZI
        Case "Standartlaştırıldı": DurumYaziHex = "1E7B4D"    '#HEX DURUM_Standartlaştırıldı_YAZI
        Case "Beklemede":          DurumYaziHex = "5A6572"    '#HEX DURUM_Beklemede_YAZI
        Case "Reddedildi":         DurumYaziHex = "9A2E2E"    '#HEX DURUM_Reddedildi_YAZI
        Case Else:                 DurumYaziHex = CLR_METIN_KOYU
    End Select
End Function


' ---------------------------------------------------------------------------
'  Oncelik sinifi renkleri
' ---------------------------------------------------------------------------
Public Function OncelikZeminHex(ByVal sinif As String) As String
    Select Case sinif
        Case "Hızlı Kazanım":      OncelikZeminHex = "D2EAD9"  '#HEX ONCELIK_Hızlı Kazanım_ZEMIN
        Case "Büyük Proje":        OncelikZeminHex = "D6E4F0"  '#HEX ONCELIK_Büyük Proje_ZEMIN
        Case "Doldurma İşi":       OncelikZeminHex = "FCEBCF"  '#HEX ONCELIK_Doldurma İşi_ZEMIN
        Case "Değerlendirme Dışı": OncelikZeminHex = "EDEFF2"  '#HEX ONCELIK_Değerlendirme Dışı_ZEMIN
        Case Else:                 OncelikZeminHex = CLR_BEYAZ
    End Select
End Function

Public Function OncelikYaziHex(ByVal sinif As String) As String
    Select Case sinif
        Case "Hızlı Kazanım":      OncelikYaziHex = "1E7B4D"   '#HEX ONCELIK_Hızlı Kazanım_YAZI
        Case "Büyük Proje":        OncelikYaziHex = "1B5FAA"   '#HEX ONCELIK_Büyük Proje_YAZI
        Case "Doldurma İşi":       OncelikYaziHex = "C77700"   '#HEX ONCELIK_Doldurma İşi_YAZI
        Case "Değerlendirme Dışı": OncelikYaziHex = "8A93A1"   '#HEX ONCELIK_Değerlendirme Dışı_YAZI
        Case Else:                 OncelikYaziHex = CLR_METIN_GRI
    End Select
End Function


' ---------------------------------------------------------------------------
'  Bir hucre araligini durum rozetine cevirir (zemin + yazi + hizalama).
' ---------------------------------------------------------------------------
Public Sub DurumRozetiUygula(ByVal hedef As Range, ByVal durum As String)
    With hedef
        .Interior.Color = HexRGB(DurumZeminHex(durum))
        .Font.Color = HexRGB(DurumYaziHex(durum))
        .Font.Name = FONT_AILE
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub

Public Sub OncelikRozetiUygula(ByVal hedef As Range, ByVal sinif As String)
    With hedef
        .Interior.Color = HexRGB(OncelikZeminHex(sinif))
        .Font.Color = HexRGB(OncelikYaziHex(sinif))
        .Font.Name = FONT_AILE
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
End Sub


' ---------------------------------------------------------------------------
'  Kaynak kodunun Turkce karakterleri dogru aktarildi mi?
'  Testler bu fonksiyonu cagirir; ".bas" dosyalari yanlis kod sayfasiyla
'  iceri aktarilmissa donen metin bozulur ve test kirmizi olur.
' ---------------------------------------------------------------------------
Public Function KodlamaSinamasi() As String
    KodlamaSinamasi = "Ölçümleniyor|Standartlaştırıldı|Hızlı Kazanım|İşi|ĞÜŞİÖÇ|ğüşıöç"
End Function
