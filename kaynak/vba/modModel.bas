Attribute VB_Name = "modModel"
Option Explicit

' ============================================================================
'  modModel -- Degerlendirme yontemi
'
'  Bu modul bir form degil, bir iyilestirme yontemi tanimlar: PDCA durum
'  akisi ve durumlarin anlamlari. Ekran kodlari bu tanimlara bagimlidir;
'  tanimlar burada tek yerde durur.
' ============================================================================

' --- Durumlar -------------------------------------------------------------
Public Const DURUM_YENI As String = "Yeni"
Public Const DURUM_DEGERLENDIRMEDE As String = "Değerlendirmede"
Public Const DURUM_PLANLANDI As String = "Planlandı"
Public Const DURUM_PILOT As String = "Pilot Uygulamada"
Public Const DURUM_OLCUM As String = "Ölçümleniyor"
Public Const DURUM_STANDART As String = "Standartlaştırıldı"
Public Const DURUM_BEKLEMEDE As String = "Beklemede"
Public Const DURUM_REDDEDILDI As String = "Reddedildi"


' ---------------------------------------------------------------------------
'  Durum listesi -- ekranlardaki dogrulama listelerinin kaynagi.
'  Sira PDCA akisini yansitir; iki yan cikis (Beklemede, Reddedildi) sonda.
' ---------------------------------------------------------------------------
Public Function Durumlar() As Variant
    Durumlar = Array(DURUM_YENI, DURUM_DEGERLENDIRMEDE, DURUM_PLANLANDI, _
                     DURUM_PILOT, DURUM_OLCUM, DURUM_STANDART, _
                     DURUM_BEKLEMEDE, DURUM_REDDEDILDI)
End Function

Public Function DurumGecerliMi(ByVal durum As String) As Boolean
    Dim v As Variant
    For Each v In Durumlar()
        If StrComp(CStr(v), durum, vbTextCompare) = 0 Then
            DurumGecerliMi = True
            Exit Function
        End If
    Next v
    DurumGecerliMi = False
End Function

' Listede mantikli siralama icin durumun akistaki sira numarasi.
Public Function DurumSirasi(ByVal durum As String) As Long
    Dim i As Long, liste As Variant
    liste = Durumlar()
    For i = LBound(liste) To UBound(liste)
        If StrComp(CStr(liste(i)), durum, vbTextCompare) = 0 Then
            DurumSirasi = i + 1
            Exit Function
        End If
    Next i
    DurumSirasi = 99
End Function


' ---------------------------------------------------------------------------
'  PDCA esleme -- her durum dongunun hangi adimina denk gelir.
' ---------------------------------------------------------------------------
Public Function PDCA(ByVal durum As String) As String
    Select Case durum
        Case DURUM_DEGERLENDIRMEDE, DURUM_PLANLANDI: PDCA = "Planla"
        Case DURUM_PILOT:                            PDCA = "Uygula"
        Case DURUM_OLCUM:                            PDCA = "Kontrol"
        Case DURUM_STANDART:                         PDCA = "Önlem"
        Case Else:                                   PDCA = "—"
    End Select
End Function


' ---------------------------------------------------------------------------
'  Uygulamaya gecmis sayilan durumlar
'
'  Panodaki "uygulamaya gecmis" sayaci bunlari sayar: oneri artik kagit
'  uzerinde degil, sahada denenmis ya da surece girmistir.
' ---------------------------------------------------------------------------
Public Function UygulanmisMi(ByVal durum As String) As Boolean
    Select Case durum
        Case DURUM_PILOT, DURUM_OLCUM, DURUM_STANDART
            UygulanmisMi = True
        Case Else
            UygulanmisMi = False
    End Select
End Function

' Henuz sonuclanmamis, Değerlendirme ekibinin ilgilenmesi gereken oneriler.
Public Function BekliyorMu(ByVal durum As String) As Boolean
    Select Case durum
        Case DURUM_YENI, DURUM_DEGERLENDIRMEDE
            BekliyorMu = True
        Case Else
            BekliyorMu = False
    End Select
End Function
