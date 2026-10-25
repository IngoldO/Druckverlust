Attribute VB_Name = "Oberflaeche"
Option Explicit

Public Sub SystemkennlinieBerechnen()
    On Error GoTo Fehler
    Dim ws As Worksheet, out As Worksheet, r As Betriebspunkt
    Dim a As Anlagendaten, qmin As Double, qmax As Double, punkte As Double, q As Double
    Dim i As Long, n As Long, daten() As Variant, meldung As String
    Set ws = ThisWorkbook.Worksheets("Eingabe")
    Set out = ThisWorkbook.Worksheets("Ergebnisse")
    a = LeseAnlage()
    qmin = Zahl(ws.Range("B13"), "Q min")
    qmax = Zahl(ws.Range("B14"), "Q max")
    punkte = Zahl(ws.Range("B15"), "Berechnungspunkte")
    If qmin < 0 Or qmax <= qmin Then Err.Raise vbObjectError + 21, , "Erforderlich: 0 <= Q min < Q max."
    If punkte < 2 Or punkte > 10000 Or punkte <> Fix(punkte) Then Err.Raise vbObjectError + 22, , "Punktanzahl muss ganzzahlig zwischen 2 und 10000 sein."
    n = CLng(punkte)
    ReDim daten(1 To n, 1 To 12)
    For i = 1 To n
        q = qmin + (qmax - qmin) * (i - 1) / (n - 1)
        r = Anlagenpunkt(q, a)
        daten(i, 1) = q: daten(i, 2) = r.V: daten(i, 3) = r.re
        daten(i, 4) = r.Lambda: daten(i, 5) = r.RohrBar: daten(i, 6) = r.EinzelBar
        daten(i, 7) = r.HoeheBar: daten(i, 8) = r.BehaelterBar: daten(i, 9) = r.GesamtBar
        daten(i, 10) = r.Regime: daten(i, 11) = r.Hinweis
        daten(i, 12) = a.ZBlende * a.Rho * r.V ^ 2 / 200000#
    Next i
    out.Range("A5:L10004").ClearContents
    out.Range("A5").Resize(n, 12).Value = daten
    out.Range("A2").Value = "Gueltiger Lauf: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | Summe n*zeta = " & Format(a.Z, "0.000")
    out.Range("B3").Value = "Armaturen enthalten aktive Lochblende; Spalte L ist darin enthalten."
    ZeichneKennlinie out, n
    out.Activate
    Exit Sub
Fehler:
    meldung = Err.Description
    On Error Resume Next
    ThisWorkbook.Worksheets("Ergebnisse").Range("A2").Value = "UNGUELTIG / NICHT AKTUELL: " & meldung
    MsgBox meldung & vbCrLf & "Berechnung abgebrochen. Vorhandene Ergebnisse sind nicht aktuell. Eingaben pruefen.", vbExclamation
End Sub

Public Sub ZeichneKennlinie(ByVal ws As Worksheet, ByVal n As Long)
    Dim ch As ChartObject, s As Series, cols As Variant, names As Variant, i As Long
    If ws.ChartObjects.Count = 0 Then
        Set ch = ws.ChartObjects.Add(ws.Range("N4").Left, ws.Range("N4").Top, 900, 420)
    Else
        Set ch = ws.ChartObjects(1)
    End If
    cols = Array(5, 6, 9)
    names = Array("Rohrreibung", "Einzelwiderstaende", "System gesamt")
    With ch.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        For i = 0 To 2
            Set s = .SeriesCollection.NewSeries
            s.Name = names(i)
            s.XValues = ws.Range("A5").Resize(n, 1)
            s.Values = ws.Cells(5, cols(i)).Resize(n, 1)
        Next i
        .HasTitle = True
        .ChartTitle.Text = "Systemkennlinie | statische Anteile im Gesamtwert"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Volumenstrom Q [m3/h]"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Differenzdruck [bar]"
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
End Sub

Public Sub Einrichtung()
    Dim ws As Worksheet, b As Shape, i As Long, namen As Variant, makros As Variant
    Set ws = ThisWorkbook.Worksheets("Eingabe")
    namen = Array("Systemkennlinie berechnen", "DN-Vorschlag uebernehmen", "Armatur hinzufuegen", "Validierung starten", "Einzelpunkt berechnen", "Pumpenbetriebspunkt", "Blenden-DN uebernehmen")
    makros = Array("SystemkennlinieBerechnen", "DNAnwenden", "ArmaturHinzufuegen", "ValidierungStarten", "EinzelpunktBerechnen", "PumpenbetriebspunktBerechnen", "BlendenDNAnwenden")
    For i = 0 To UBound(namen)
        On Error Resume Next
        ws.Shapes("dvButton" & i).Delete
        On Error GoTo 0
        Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("F5").Left, ws.Range("F5").Top + i * 45, 250, 34)
        b.Name = "dvButton" & i
        b.TextFrame.Characters.Text = namen(i)
        b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & makros(i)
        b.Fill.ForeColor.RGB = RGB(24, 84, 105)
        b.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    Next i
    MsgBox "Einrichtung abgeschlossen. DN-Vorschlag bei Bedarf uebernehmen; Berechnung ueber Schaltflaeche starten.", vbInformation
End Sub
