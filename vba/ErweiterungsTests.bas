Attribute VB_Name = "ErweiterungsTests"
Option Explicit

Private Sub Nah(ByVal actual As Double, ByVal expected As Double, ByVal label As String)
    If Abs(actual - expected) > 0.000001 * (1 + Abs(expected)) Then Err.Raise vbObjectError + 150, , label & ": Ist=" & actual & " Soll=" & expected
End Sub

Public Sub ErweiterungenValidieren()
    On Error GoTo Fehler
    Dim a As Anlagendaten, r As Betriebspunkt, qs(1 To 3) As Double, hs(1 To 3) As Double
    Dim count As Long, i As Long, code As Long, wert As Double
    Nah BlendenZeta(0.05#, 0.010000000000000002#, 0.55000000000000004#), 1973.9515678798182#, "Blende beta=0.2 Cd=0.55"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.010000000000000002#, 0.60999999999999999#), 1597.0425278416976#, "Blende beta=0.2 Cd=0.61"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.010000000000000002#, 0.69999999999999996#), 1204.0699651312425#, "Blende beta=0.2 Cd=0.7"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.012500000000000001#, 0.55000000000000004#), 786.8727038561268#, "Blende beta=0.25 Cd=0.55"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.012500000000000001#, 0.60999999999999999#), 634.90500775151759#, "Blende beta=0.25 Cd=0.61"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.012500000000000001#, 0.69999999999999996#), 476.73943596684489#, "Blende beta=0.25 Cd=0.7"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.025000000000000001#, 0.55000000000000004#), 37.361900500181477#, "Blende beta=0.5 Cd=0.55"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.025000000000000001#, 0.60999999999999999#), 29.456901619036948#, "Blende beta=0.5 Cd=0.61"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.025000000000000001#, 0.69999999999999996#), 21.367291381299509#, "Blende beta=0.5 Cd=0.7"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.037500000000000006#, 0.55000000000000004#), 3.4352453681530157#, "Blende beta=0.75 Cd=0.55"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.037500000000000006#, 0.60999999999999999#), 2.5884723077873431#, "Blende beta=0.75 Cd=0.61"
    count = count + 1
    Nah BlendenZeta(0.05#, 0.037500000000000006#, 0.69999999999999996#), 1.7576461319778298#, "Blende beta=0.75 Cd=0.7"
    count = count + 1
    a.D = 0.05: a.Rho = 1000: a.Mu = 0.001
    a.BlendeD = 0.05: a.ZBlende = BlendenZeta(0.05, 0.025, 0.61): a.Z = 2 + a.ZBlende
    r = Anlagenpunkt(10, a)
    Nah r.GesamtBar, 0.31479017340853288#, "Anlagenpunkt inkl. Blende"
    count = count + 1
    r = Anlagenpunkt(0, a): Nah r.GesamtBar, 0, "Blende Q=0"
    count = count + 1
    a.ZBlende = 0: a.BlendeD = 0: a.Z = 2
    qs(1) = 0: qs(2) = 20: qs(3) = 40
    hs(1) = 4 / (a.Rho * G / 100000#): hs(2) = hs(1) / 2: hs(3) = 0
    r = PumpenSchnittpunkt(a, qs, hs)
    Nah r.Q, 37.226436301264144#, "Analytischer Pumpenschnittpunkt"
    count = count + 1
    Nah Pumpendruck(r.Q, qs, hs, a.Rho), r.GesamtBar, "Druckresiduum"
    count = count + 1
    a.Z = 0: a.DP = 1
    hs(1) = 30: hs(2) = 20: hs(3) = 0: qs(2) = 10: qs(3) = 20
    r = PumpenSchnittpunkt(a, qs, hs)
    Nah r.Q, 14.901418935110359#, "Stueckweise lineare Pumpe"
    count = count + 1
    a.DP = 1000 * G * hs(1) / 100000#: r = PumpenSchnittpunkt(a, qs, hs)
    Nah r.Q, 0, "Linker Randpunkt"
    count = count + 1
    a.DP = 0: r = PumpenSchnittpunkt(a, qs, hs)
    Nah r.Q, 20, "Rechter Randpunkt"
    count = count + 1
    For i = 1 To 10
        On Error Resume Next
        Err.Clear
        Select Case i
        Case 1: wert = BlendenZeta(0.05, 0.05, 0.61)
        Case 2: wert = BlendenZeta(0.05, 0.001, 0.61)
        Case 3: wert = BlendenZeta(0.05, 0.025, 0)
        Case 4: wert = Pumpendruck(-1, qs, hs, 1000)
        Case 5: wert = Pumpendruck(21, qs, hs, 1000)
        Case 6: a.DP = 100: r = PumpenSchnittpunkt(a, qs, hs)
        Case 7: a.DP = -1: r = PumpenSchnittpunkt(a, qs, hs)
        Case 8: qs(2) = 0: wert = Pumpendruck(1, qs, hs, 1000)
        Case 9: qs(2) = 10: hs(2) = 40: wert = Pumpendruck(1, qs, hs, 1000)
        Case 10:
            a.L = 100000: a.DP = 0: a.Z = 0
            qs(1) = 0.3: qs(2) = 0.32515483964654363#: qs(3) = 0.35
            hs(1) = 8.1134481313394016#
            hs(2) = 8.103250969209622#
            hs(3) = 8.0930538070798423#
            r = PumpenSchnittpunkt(a, qs, hs)
        End Select
        code = Err.Number
        On Error GoTo Fehler
        If code = 0 Then Err.Raise vbObjectError + 151, , "Fehlerfall akzeptiert: " & i
        count = count + 1
    Next i
    MsgBox count & " Erweiterungs-Testfaelle bestanden.", vbInformation
    Exit Sub
Fehler:
    MsgBox "ERWEITERUNGS-VALIDIERUNG FEHLGESCHLAGEN: " & Err.Description, vbCritical
    Err.Raise vbObjectError + 152, "ErweiterungenValidieren", "Validierung fehlgeschlagen"
End Sub
