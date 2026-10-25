Attribute VB_Name = "EngineeringBericht"
Option Explicit

Public Sub BerichtsSchaltflaeche()
    Dim ws As Worksheet, b As Shape
    Set ws = ThisWorkbook.Worksheets("Systembericht")
    On Error Resume Next
    ws.Shapes("dvBerichtButton").Delete
    On Error GoTo 0
    Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("J3").Left, ws.Range("J3").Top, 220, 32)
    b.Name = "dvBerichtButton"
    b.TextFrame.Characters.Text = "Systembericht erstellen"
    b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!SystemberichtErstellen"
End Sub

Private Sub Anhang(ByVal ws As Worksheet, ByRef row As Long, ByVal txt As String)
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 8)).Merge
    ws.Cells(row, 1).Value = txt
    ws.Cells(row, 1).WrapText = True
    ws.Cells(row, 1).Font.Size = 10
    ws.Rows(row).RowHeight = 14 * (2 + Int(Len(txt) / 85))
    row = row + 1
End Sub

Public Sub SystemberichtErstellen()
    On Error GoTo Fehler
    Dim ws As Worksheet, inp As Worksheet, a As Anlagendaten, r As Betriebspunkt, op As Betriebspunkt
    Dim qmin As Double, qmax As Double, nval As Double, n As Long, i As Long, row As Long
    Dim q As Double, mit As Double, qs() As Double, hs() As Double, daten() As Variant
    Dim ch As ChartObject, s As Series, meldung As String, warnungen As Long
    Dim hinweise As Collection, hinweis As Variant
    Dim t As ListObject, lr As ListRow, typ As String, z As Double, sheet As Variant, j As Long
    Set ws = ThisWorkbook.Worksheets("Systembericht")
    Set inp = ThisWorkbook.Worksheets("Eingabe")
    ws.Range("A6").Value = "NICHT AKTUELL: Bericht wird neu berechnet"
    a = LeseAnlage()
    qmin = Zahl(inp.Range("B13"), "Q min")
    qmax = Zahl(inp.Range("B14"), "Q max")
    nval = Zahl(inp.Range("B15"), "Punkte")
    If qmin < 0 Or qmax <= qmin Or nval < 2 Or nval > 10000 Or nval <> Fix(nval) Then Err.Raise vbObjectError + 70, , "Kennlinienbereich/Punktanzahl ungueltig."
    n = CLng(nval)
    mit = Zahl(ws.Range("D5"), "Pumpe einschliessen")
    If mit <> 0 And mit <> 1 Then Err.Raise vbObjectError + 71, , "Pumpe einschliessen muss 0 oder 1 sein."
    If mit = 1 Then
        LesePumpe qs, hs
        op = PumpenSchnittpunkt(a, qs, hs)
        If qs(LBound(qs)) < qmin Then qmin = qs(LBound(qs))
        If qs(UBound(qs)) > qmax Then qmax = qs(UBound(qs))
    End If
    Set hinweise = New Collection
    ReDim daten(1 To n, 1 To 4)
    For i = 1 To n
        q = qmin + (qmax - qmin) * (i - 1) / (n - 1)
        r = Anlagenpunkt(q, a)
        daten(i, 1) = q: daten(i, 2) = r.RohrBar: daten(i, 3) = r.EinzelBar: daten(i, 4) = r.GesamtBar
        If Len(r.Hinweis) > 0 Then
            warnungen = warnungen + 1
            On Error Resume Next
            hinweise.Add r.Hinweis, r.Hinweis
            On Error GoTo Fehler
        End If
    Next i
    ws.Range("N5:U10004").ClearContents
    ws.Range("N5").Resize(n, 4).Value = daten
    ws.Range("A8").Value = "D=" & Format(a.D * 1000, "0.###") & " mm; L=" & a.L & " m; Rauigkeit=" & a.Eps * 1000 & " mm; Hoehe=" & a.H & " m" & vbLf & "rho=" & a.Rho & " kg/m3; mu=" & a.Mu * 1000 & " mPa s; Behaelterdruck=" & a.DP & " bar" & vbLf & "Q=" & Format(qmin, "0.###") & " bis " & Format(qmax, "0.###") & " m3/h; " & n & " Punkte; Zeta gesamt=" & Format(a.Z, "0.###")
    If mit = 1 Then
        ws.Range("A11").Value = "Pumpenbetriebspunkt: Q=" & Format(op.Q, "0.0000") & " m3/h; dp=" & Format(op.GesamtBar, "0.000000") & " bar" & vbLf & "Druckresiduum=" & Format(Pumpendruck(op.Q, qs, hs, a.Rho) - op.GesamtBar, "0.000E+00") & " bar. Hinweise und Modellgrenzen im Anhang."
        ws.Range("T5").Value = op.Q: ws.Range("U5").Value = op.GesamtBar
        For i = LBound(qs) To UBound(qs)
            ws.Cells(5 + i - LBound(qs), 18).Value = qs(i)
            ws.Cells(5 + i - LBound(qs), 19).Value = a.Rho * G * hs(i) / 100000#
        Next i
    Else
        ws.Range("A11").Value = "Systemkennlinie ohne Pumpe; kein Pumpenbetriebspunkt berechnet. Warnungen in " & warnungen & " von " & n & " Punkten; siehe Anhang."
    End If
    On Error Resume Next
    ws.ChartObjects("dvBerichtChart").Delete
    On Error GoTo Fehler
    Set ch = ws.ChartObjects.Add(ws.Range("A14").Left, ws.Range("A14").Top, 600, 300)
    ch.Name = "dvBerichtChart"
    With ch.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .PlotVisibleOnly = False
        For i = 15 To 17
            Set s = .SeriesCollection.NewSeries
            If i = 15 Then s.Name = "Rohrreibung"
            If i = 16 Then s.Name = "Einzel inkl. Blende/Kv"
            If i = 17 Then s.Name = "System gesamt"
            s.XValues = ws.Range("N5").Resize(n, 1): s.Values = ws.Cells(5, i).Resize(n, 1)
        Next i
        If mit = 1 Then
            Set s = .SeriesCollection.NewSeries
            s.Name = "Pumpe": s.XValues = ws.Range("R5").Resize(UBound(qs), 1): s.Values = ws.Range("S5").Resize(UBound(qs), 1)
            Set s = .SeriesCollection.NewSeries
            s.Name = "Betriebspunkt": s.XValues = ws.Range("T5"): s.Values = ws.Range("U5")
            s.ChartType = xlXYScatter: s.MarkerStyle = xlMarkerStyleCircle: s.MarkerSize = 8
        End If
        .HasTitle = False
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Q [m3/h]"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "Differenzdruck [bar]"
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
    End With
    row = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If row >= 47 Then ws.Range("A47:H" & row).UnMerge: ws.Range("A47:H" & row).ClearContents
    row = 47
    Anhang ws, row, "Berechnungsstand: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "; Modellversion 1.2; Warnungen in " & warnungen & " von " & n & " Kennlinienpunkten."
    For i = 5 To 16
        Anhang ws, row, CStr(inp.Cells(i, 1).Value) & ": " & CStr(inp.Cells(i, 2).Value) & " " & CStr(inp.Cells(i, 3).Value)
    Next i
    For Each sheet In Array("Lochblende", "KvArmatur")
        j = 10: If sheet = "KvArmatur" Then j = 7
        For i = 5 To j
            Anhang ws, row, CStr(sheet) & " | " & CStr(ThisWorkbook.Worksheets(sheet).Cells(i, 1).Value) & ": " & CStr(ThisWorkbook.Worksheets(sheet).Cells(i, 2).Value)
        Next i
    Next sheet
    Set t = inp.ListObjects("tblArmaturen")
    For Each lr In t.ListRows
        typ = CStr(lr.Range.Cells(1, 1).Value)
        If Len(typ) > 0 Then
            z = StandardZeta(typ)
            If Not IsEmpty(lr.Range.Cells(1, 3).Value) Then z = Zahl(lr.Range.Cells(1, 3), "Zeta")
            Anhang ws, row, typ & ": Anzahl=" & lr.Range.Cells(1, 2).Value & "; wirksames Zeta=" & z
        End If
    Next lr
    If mit = 1 Then
        For i = LBound(qs) To UBound(qs)
            Anhang ws, row, "Pumpenpunkt " & i & ": Q=" & qs(i) & " m3/h; H=" & hs(i) & " m"
        Next i
        Anhang ws, row, "Betriebspunkt: " & op.Regime & "; " & op.Hinweis
    End If
    Anhang ws, row, "Modell: Zeta/Cd/Kv konstant; Q=0 nur statisch; Re 2300-4000 unsicher. Kein Nachweis fuer Kavitation/NPSH/Festigkeit."
    For Each hinweis In hinweise
        Anhang ws, row, "Warnung in Kennlinie: " & CStr(hinweis)
    Next hinweis
    ws.PageSetup.PrintArea = "$A$1:$H$" & (row - 1)
    ws.PageSetup.Zoom = 95
    ws.ResetAllPageBreaks
    ws.HPageBreaks.Add Before:=ws.Range("A45")
    ws.Range("A6").Value = "BERECHNET: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & " | Nach Aenderungen Bericht neu erstellen."
    ws.Activate
    Exit Sub
Fehler:
    meldung = Err.Description
    On Error Resume Next
    ws.Range("A6").Value = "UNGUELTIG / NICHT AKTUELL: " & meldung
    MsgBox meldung, vbExclamation, "Systembericht"
End Sub
