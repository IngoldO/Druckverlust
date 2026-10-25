Attribute VB_Name = "Pumpenmodell"
Option Explicit

Public Function Anlagenpunkt(ByVal q As Double, ByRef a As Anlagendaten) As Betriebspunkt
    Dim r As Betriebspunkt, reLokal As Double
    r = BerechnePunkt(q, a.D, a.L, a.Eps, a.Rho, a.Mu, a.H, a.Z, a.DP)
    If a.ZBlende > 0 Then
        r.Hinweis = r.Hinweis & " Lochblende: Cd-Naeherung; bleibender Verlust, kein Messdifferenzdruck."
        reLokal = a.Rho * (q / 3600#) * 4 / (4 * Atn(1#) * a.BlendeD * a.Mu)
        If q > 0 And reLokal < 10000 Then r.Hinweis = r.Hinweis & " Blenden-Re <10000: Modell besonders unsicher."
        If Abs(a.BlendeD - a.D) > 0.000000001 Then r.Hinweis = r.Hinweis & " Abweichender Blenden-D: Reduzierungen separat beruecksichtigen."
    End If
    If a.ZKv > 0 Then r.Hinweis = r.Hinweis & " Kv: konstante Oeffnung; turbulent, nicht kavitierend angenommen; Herstellergrenzen pruefen."
    Anlagenpunkt = r
End Function

Public Sub PruefePumpendaten(ByRef qs() As Double, ByRef hs() As Double, ByVal rho As Double)
    Dim i As Long, first As Long, last As Long
    first = LBound(qs): last = UBound(qs)
    If last <= first Or LBound(hs) <> first Or UBound(hs) <> last Or rho <= 0 Then Err.Raise vbObjectError + 40, , "Mindestens zwei Pumpenpunkte und positive Dichte erforderlich."
    For i = first To last
        If qs(i) < 0 Or hs(i) < 0 Then Err.Raise vbObjectError + 41, , "Pumpen-Q und Foerderhoehe muessen >=0 sein."
        If i > first Then
            If qs(i) <= qs(i - 1) Then Err.Raise vbObjectError + 42, , "Pumpen-Q muss streng aufsteigend sein."
            If hs(i) >= hs(i - 1) Then Err.Raise vbObjectError + 43, , "Foerderhoehe muss streng fallen; nichtmonotone Kennlinien werden nicht unterstuetzt."
        End If
    Next i
End Sub

Private Function InterpolierterDruck(ByVal q As Double, ByRef qs() As Double, ByRef hs() As Double, ByVal rho As Double) As Double
    Dim i As Long, hoehe As Double
    If q < qs(LBound(qs)) Or q > qs(UBound(qs)) Then Err.Raise vbObjectError + 44, , "Q ausserhalb Pumpentabelle; keine Extrapolation."
    For i = LBound(qs) + 1 To UBound(qs)
        If q <= qs(i) Then
            hoehe = hs(i - 1) + (hs(i) - hs(i - 1)) * (q - qs(i - 1)) / (qs(i) - qs(i - 1))
            InterpolierterDruck = rho * G * hoehe / 100000#
            Exit Function
        End If
    Next i
End Function

Public Function Pumpendruck(ByVal q As Double, ByRef qs() As Double, ByRef hs() As Double, ByVal rho As Double) As Double
    PruefePumpendaten qs, hs, rho
    Pumpendruck = InterpolierterDruck(q, qs, hs, rho)
End Function

Public Function PumpenSchnittpunkt(ByRef a As Anlagendaten, ByRef qs() As Double, ByRef hs() As Double) As Betriebspunkt
    Dim lo As Double, hi As Double, mid As Double, f As Double, flo As Double, fhi As Double
    Dim i As Long, r As Betriebspunkt
    PruefePumpendaten qs, hs, a.Rho
    lo = qs(LBound(qs)): hi = qs(UBound(qs))
    r = Anlagenpunkt(lo, a): flo = InterpolierterDruck(lo, qs, hs, a.Rho) - r.GesamtBar
    If Abs(flo) <= 0.00000001 Then
        PumpenSchnittpunkt = r
        Exit Function
    End If
    r = Anlagenpunkt(hi, a): fhi = InterpolierterDruck(hi, qs, hs, a.Rho) - r.GesamtBar
    If Abs(fhi) <= 0.00000001 Then
        PumpenSchnittpunkt = r
        Exit Function
    End If
    If flo < 0 Or fhi > 0 Then Err.Raise vbObjectError + 45, , "Kein Schnittpunkt im angegebenen Pumpenbereich; keine Extrapolation."
    ' Decreasing pump curve and increasing system characteristic: unique root.
    For i = 1 To 120
        mid = lo + (hi - lo) / 2
        r = Anlagenpunkt(mid, a)
        f = InterpolierterDruck(mid, qs, hs, a.Rho) - r.GesamtBar
        If Abs(f) <= 0.00000001 Then
            PumpenSchnittpunkt = r
            Exit Function
        End If
        If f > 0 Then lo = mid Else hi = mid
    Next i
    ' In particular, do not report a sign change across the Re=2300 jump as a root.
    Err.Raise vbObjectError + 46, , "Kein konvergierter Schnittpunkt; moeglicher Sprung im Uebergangsmodell."
End Function
