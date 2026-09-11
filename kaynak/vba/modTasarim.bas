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
' Marka
Public Const CLR_ANA_LACIVERT As String = "002D62"    '#HEX ANA_LACIVERT
Public Const CLR_DERIN_LACIVERT As String = "001F44"  '#HEX DERIN_LACIVERT
Public Const CLR_KURUMSAL_MAVI As String = "1F6FEB"   '#HEX KURUMSAL_MAVI
Public Const CLR_ACIK_MAVI As String = "E8F1FD"       '#HEX ACIK_MAVI
Public Const CLR_SOLUK_MAVI As String = "A8C0E0"      '#HEX SOLUK_MAVI

' Notr basamaklar
Public Const CLR_BUZ_ZEMIN As String = "F5F7FA"       '#HEX BUZ_ZEMIN
Public Const CLR_BEYAZ As String = "FFFFFF"           '#HEX BEYAZ
Public Const CLR_CIZGI_INCE As String = "EEF2F7"      '#HEX CIZGI_INCE
Public Const CLR_CIZGI_GRI As String = "E6EBF2"       '#HEX CIZGI_GRI
Public Const CLR_GOLGE_GRI As String = "DCE3EC"       '#HEX GOLGE_GRI
Public Const CLR_ALAN_CIZGI As String = "C9D4E3"      '#HEX ALAN_CIZGI
Public Const CLR_METIN_KOYU As String = "0F1B2E"      '#HEX METIN_KOYU
Public Const CLR_METIN_GRI As String = "64748B"       '#HEX METIN_GRI
Public Const CLR_METIN_SOLUK As String = "94A3B8"     '#HEX METIN_SOLUK

' Durum bantlari ve vurgular
Public Const CLR_ONAY_ZEMIN As String = "E3F5EA"      '#HEX ONAY_ZEMIN
Public Const CLR_ONAY_YAZI As String = "15803D"       '#HEX ONAY_YAZI
Public Const CLR_UYARI_ZEMIN As String = "FDE8E8"     '#HEX UYARI_ZEMIN
Public Const CLR_UYARI_YAZI As String = "B42318"      '#HEX UYARI_YAZI
Public Const CLR_VURGU_KEHRIBAR As String = "B45309"  '#HEX VURGU_KEHRIBAR

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
'  Bant isaretleri -- onay (U+2713) ve uyari (U+26A0)
'
'  VBA kaynagi ANSI kod sayfasinda saklanir (Turkce Windows'ta cp1254).
'  Turkce harfler o sayfada vardir, bu iki isaret YOKTUR: kaynaga duz
'  yazildiklarinda hucreye "?" olarak iner -- ekranlarda boyle goruldu.
'  Bu yuzden karakter kodundan uretilirler. test_uretim.py kod satirlarinda
'  cp1254 disi karakter kalmadigini denetler.
' ---------------------------------------------------------------------------
Public Function IsaretOnay() As String
    IsaretOnay = ChrW$(&H2713)
End Function

Public Function IsaretUyari() As String
    IsaretUyari = ChrW$(&H26A0)
End Function


' ---------------------------------------------------------------------------
'  Durum renkleri
' ---------------------------------------------------------------------------
Public Function DurumZeminHex(ByVal durum As String) As String
    Select Case durum
        Case "Yeni":               DurumZeminHex = "E8F1FD"   '#HEX DURUM_Yeni_ZEMIN
        Case "Değerlendirmede":    DurumZeminHex = "FEF3C7"   '#HEX DURUM_Değerlendirmede_ZEMIN
        Case "Planlandı":          DurumZeminHex = "EDE9FE"   '#HEX DURUM_Planlandı_ZEMIN
        Case "Pilot Uygulamada":   DurumZeminHex = "CFFAFE"   '#HEX DURUM_Pilot Uygulamada_ZEMIN
        Case "Ölçümleniyor":       DurumZeminHex = "E7F6EC"   '#HEX DURUM_Ölçümleniyor_ZEMIN
        Case "Standartlaştırıldı": DurumZeminHex = "DCFCE7"   '#HEX DURUM_Standartlaştırıldı_ZEMIN
        Case "Beklemede":          DurumZeminHex = "F1F5F9"   '#HEX DURUM_Beklemede_ZEMIN
        Case "Reddedildi":         DurumZeminHex = "FDE8E8"   '#HEX DURUM_Reddedildi_ZEMIN
        Case Else:                 DurumZeminHex = CLR_BEYAZ
    End Select
End Function

Public Function DurumYaziHex(ByVal durum As String) As String
    Select Case durum
        Case "Yeni":               DurumYaziHex = "1D4ED8"    '#HEX DURUM_Yeni_YAZI
        Case "Değerlendirmede":    DurumYaziHex = "92400E"    '#HEX DURUM_Değerlendirmede_YAZI
        Case "Planlandı":          DurumYaziHex = "5B21B6"    '#HEX DURUM_Planlandı_YAZI
        Case "Pilot Uygulamada":   DurumYaziHex = "0E7490"    '#HEX DURUM_Pilot Uygulamada_YAZI
        Case "Ölçümleniyor":       DurumYaziHex = "3F7A47"    '#HEX DURUM_Ölçümleniyor_YAZI
        Case "Standartlaştırıldı": DurumYaziHex = "15803D"    '#HEX DURUM_Standartlaştırıldı_YAZI
        Case "Beklemede":          DurumYaziHex = "64748B"    '#HEX DURUM_Beklemede_YAZI
        Case "Reddedildi":         DurumYaziHex = "B42318"    '#HEX DURUM_Reddedildi_YAZI
        Case Else:                 DurumYaziHex = CLR_METIN_KOYU
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



' ---------------------------------------------------------------------------
'  Kaynak kodunun Turkce karakterleri dogru aktarildi mi?
'  Testler bu fonksiyonu cagirir; ".bas" dosyalari yanlis kod sayfasiyla
'  iceri aktarilmissa donen metin bozulur ve test kirmizi olur.
' ---------------------------------------------------------------------------
Public Function KodlamaSinamasi() As String
    KodlamaSinamasi = "Ölçümleniyor|Standartlaştırıldı|Hızlı Kazanım|İşi|ĞÜŞİÖÇ|ğüşıöç"
End Function
