Attribute VB_Name = "Betriebsoberflaeche"
Option Explicit

Private Sub SchreibePunkt(ByVal ws As Worksheet, ByVal zeile As Long, ByRef r As Betriebspunkt, ByRef a As Anlagendaten)
    Dim data(1 To 1, 1 To 13) As Variant
    data(1, 1) = r.Q: data(1, 2) = r.V: data(1, 3) = r.Re
    data(1, 4) = r.Lambda: data(1, 5) = r.RohrBar: data(1, 6) = r.EinzelBar
    data(1, 7) = r.HoeheBar: data(1, 8) = r.BehaelterBar: data(1, 9) = r.GesamtBar
    data(1, 10) = r.Regime: data(1, 11) = r.Hinweis
    data(1, 12) = a.ZBlende * a.Rho * r.V ^ 2 / 200000#
    data(1, 13) = a.ZKv * a.Rho * r.V ^ 2 / 200000#
    ws.Cells(zeile, 1).Resize(1, 13).Value = data
End Sub

Public Sub EinzelpunktBerechnen()
    On Error GoTo Fehler
    Dim a As Anlagendaten, r As Betriebspunkt, q As Double, ws As Worksheet, meldung As String
    Set ws = ThisWorkbook.Worksheets("Betriebspunkte")
    ws.Range("A3").Value = "NICHT AKTUELL: Berechnung laeuft"
    a = LeseAnlage()
    q = Zahl(ThisWorkbook.Worksheets("Eingabe").Range("G3"), "Einzelpunkt-Q")
    r = Anlagenpunkt(q, a)
    SchreibePunkt ws, 5, r, a
    ws.Range("A3").Value = "Einzelpunkt berechnet: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Activate
    Exit Sub
Fehler:
    meldung = Err.Description
    On Error Resume Next
    ws.Range("A3").Value = "UNGUELTIG / NICHT AKTUELL: " & meldung
    MsgBox meldung, vbExclamation, "Einzelpunkt"
End Sub

Public Sub PumpenbetriebspunktBerechnen()
    On Error GoTo Fehler
    Dim a As Anlagendaten, r As Betriebspunkt, ws As Worksheet, p As Worksheet, meldung As String
    Dim qs() As Double, hs() As Double, i As Long, q As Double, data(1 To 201, 1 To 2) As Variant
    Set ws = ThisWorkbook.Worksheets("Betriebspunkte")
    Set p = ThisWorkbook.Worksheets("Pumpe")
    ws.Range("A9").Value = "NICHT AKTUELL: Berechnung laeuft"
    p.Range("A3").Value = "NICHT AKTUELL: Berechnung laeuft"
    a = LeseAnlage()
    LesePumpe qs, hs
    r = PumpenSchnittpunkt(a, qs, hs)
    SchreibePunkt ws, 11, r, a
    ws.Range("N10:O11").ClearContents
    ws.Range("N10").Value = "Pumpendruck [bar]": ws.Range("O10").Value = "Residuum [bar]"
    ws.Range("N11").Value = Pumpendruck(r.Q, qs, hs, a.Rho)
    ws.Range("O11").Value = ws.Range("N11").Value - r.GesamtBar
    For i = LBound(qs) To UBound(qs)
        p.Cells(20 + i - LBound(qs) + 1, 3).Value = a.Rho * G * hs(i) / 100000#
    Next i
    For i = 1 To 201
        q = qs(LBound(qs)) + (qs(UBound(qs)) - qs(LBound(qs))) * (i - 1) / 200#
        r = Anlagenpunkt(q, a)
        data(i, 1) = q: data(i, 2) = r.GesamtBar
    Next i
    p.Range("N5").Resize(201, 2).Value = data
    ZeichnePumpe p, UBound(qs) - LBound(qs) + 1
    ws.Range("A9").Value = "Pumpenschnittpunkt berechnet: " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    p.Range("A3").Value = ws.Range("A9").Value
    ws.Activate
    Exit Sub
Fehler:
    meldung = Err.Description
    On Error Resume Next
    ws.Range("A9").Value = "UNGUELTIG / NICHT AKTUELL: " & meldung
    p.Range("A3").Value = "UNGUELTIG / NICHT AKTUELL: " & meldung
    MsgBox meldung, vbExclamation, "Pumpenbetriebspunkt"
End Sub

Private Sub ZeichnePumpe(ByVal ws As Worksheet, ByVal n As Long)
    Dim ch As ChartObject, s As Series
    If ws.ChartObjects.Count = 0 Then
        Set ch = ws.ChartObjects.Add(ws.Range("E20").Left, ws.Range("E20").Top, 850, 440)
    Else
        Set ch = ws.ChartObjects(1)
    End If
    With ch.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Set s = .SeriesCollection.NewSeries
        s.Name = "Pumpe"
        s.XValues = ws.Range("A21").Resize(n, 1): s.Values = ws.Range("C21").Resize(n, 1)
        Set s = .SeriesCollection.NewSeries
        s.Name = "System inkl. Blende/Kv"
        s.XValues = ws.Range("N5:N205"): s.Values = ws.Range("O5:O205")
        Set s = .SeriesCollection.NewSeries
        s.Name = "Betriebspunkt"
        s.XValues = ThisWorkbook.Worksheets("Betriebspunkte").Range("A11")
        s.Values = ThisWorkbook.Worksheets("Betriebspunkte").Range("I11")
        s.ChartType = xlXYScatter
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 9
        .HasTitle = True: .ChartTitle.Text = "Pumpe und System | Betriebspunkt"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Q [m3/h]"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "Differenzdruck [bar]"
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
    End With
End Sub
