Attribute VB_Name = "modModel"
Option Explicit

' ============================================================================
'  modModel -- Kaizen yontemi
'
'  Bu modul bir form degil, bir iyilestirme yontemi tanimlar: PDCA durum
'  akisi, Toyota'nin yedi israfi (muda) ve etki/efor oncelik matrisi.
'  Ekran kodlari bu tanimlara bagimlidir; tanimlar burada tek yerde durur.
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

' --- Oncelik siniflari ----------------------------------------------------
Public Const ONCELIK_HIZLI As String = "Hızlı Kazanım"
Public Const ONCELIK_BUYUK As String = "Büyük Proje"
Public Const ONCELIK_DOLDURMA As String = "Doldurma İşi"
Public Const ONCELIK_DISI As String = "Değerlendirme Dışı"


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

' Konsolda mantikli siralama icin durumun akistaki sira numarasi.
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
'  Oncelik matrisi
'
'                 | Düşük efor (<=2)   | Yüksek efor (>=3)
'  Yüksek etki    | Hızlı Kazanım      | Büyük Proje
'  (>=3)          | önce bunlar        | planlama ve kaynak gerektirir
'  ---------------+--------------------+----------------------------
'  Düşük etki     | Doldurma İşi       | Değerlendirme Dışı
'  (<=2)          | boş kapasiteyle    | bu haliyle önerilmez
'
'  Puanlanmamis (0) kayitlar sinif disidir: henuz karar verilmemistir.
' ---------------------------------------------------------------------------
Public Function OncelikSinifi(ByVal etki As Long, ByVal efor As Long) As String
    If etki <= 0 Or efor <= 0 Then
        OncelikSinifi = ""
        Exit Function
    End If

    If etki >= modAyar.ESIK_YUKSEK_ETKI Then
        If efor <= modAyar.ESIK_DUSUK_EFOR Then
            OncelikSinifi = ONCELIK_HIZLI
        Else
            OncelikSinifi = ONCELIK_BUYUK
        End If
    Else
        If efor <= modAyar.ESIK_DUSUK_EFOR Then
            OncelikSinifi = ONCELIK_DOLDURMA
        Else
            OncelikSinifi = ONCELIK_DISI
        End If
    End If
End Function

Public Function OncelikSiniflari() As Variant
    OncelikSiniflari = Array(ONCELIK_HIZLI, ONCELIK_BUYUK, _
                             ONCELIK_DOLDURMA, ONCELIK_DISI)
End Function


' ---------------------------------------------------------------------------
'  Tasarruf kurali
'
'  Bir onerinin yillik saat/TL kazanci gostergelere YALNIZCA fayda gerceklesmis
'  sayilan durumlarda katilir. Heniz uygulanmamis bir oneri tasarruf degildir;
'  aksi halde pano gercekte olmayan bir kazanci raporlar.
' ---------------------------------------------------------------------------
Public Function TasarrufSayilirMi(ByVal durum As String) As Boolean
    Select Case durum
        Case DURUM_PILOT, DURUM_OLCUM, DURUM_STANDART
            TasarrufSayilirMi = True
        Case Else
            TasarrufSayilirMi = False
    End Select
End Function

' Henuz sonuclanmamis, Kaizen ekibinin ilgilenmesi gereken oneriler.
Public Function BekliyorMu(ByVal durum As String) As Boolean
    Select Case durum
        Case DURUM_YENI, DURUM_DEGERLENDIRMEDE
            BekliyorMu = True
        Case Else
            BekliyorMu = False
    End Select
End Function

' Kabul edilmis sayilan durumlar (kabul orani hesabinda kullanilir).
Public Function KabulEdildiMi(ByVal durum As String) As Boolean
    Select Case durum
        Case DURUM_PLANLANDI, DURUM_PILOT, DURUM_OLCUM, DURUM_STANDART
            KabulEdildiMi = True
        Case Else
            KabulEdildiMi = False
    End Select
End Function

' Karara baglanmis (artik beklemeyen) oneriler.
Public Function SonuclandiMi(ByVal durum As String) As Boolean
    SonuclandiMi = Not BekliyorMu(durum)
End Function


' ---------------------------------------------------------------------------
'  Puan listesi (etki / efor): 1-5
' ---------------------------------------------------------------------------
Public Function Puanlar() As Variant
    Puanlar = Array(1, 2, 3, 4, 5)
End Function
