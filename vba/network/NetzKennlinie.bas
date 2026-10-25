Attribute VB_Name = "NetzKennlinie"
Option Explicit

Public Sub NetzKennlinienPunkt(ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByVal einspeisung As Long, ByVal q As Double, ByVal rho As Double, ByVal mu As Double, ByRef pq() As Double, ByRef ph() As Double, ByRef aus As NetzResultat, ByRef druck As Double, Optional ByVal maxIter As Long = 100, Optional ByVal tq As Double = 0.000001, Optional ByVal tp As Double = 0.00000001)
    Dim kc() As NetzKnoten, i As Long, pump As Long, fest As Long, leer As NetzResultat
    aus = leer: druck = 0
    If q < 0 Or einspeisung < 1 Or einspeisung > UBound(k) Then NetzFehler "Kennlinie: Q>=0 und gueltigen Einspeiseknoten angeben."
    ReDim kc(1 To UBound(k))
    For i = 1 To UBound(k)
        kc(i) = k(i)
    Next i
    For i = 1 To UBound(e)
        If e(i).Aktiv And e(i).Pumpe Then pump = i
    Next i
    If pump > 0 Then
        If einspeisung <> e(pump).Nach Then NetzFehler "Mit Pumpe den Druckstutzen (Endknoten der Pumpenleitung) als Einspeisung waehlen. Alternative: Netzwerkbetriebspunkt berechnen."
        If e(pump).Rueckschlag And q = 0 Then NetzFehler "Pumpenkennlinie bei Q=0 mit Rueckschlagventil nicht eindeutig; Qmin>0 oder Einzelbetriebspunkt waehlen."
        NetzLoesen kc, e, rho, mu, pq, ph, aus, maxIter, tq, tp, q
        If aus.Erfolg Then druck = aus.PumpenDruck
    Else
        kc(einspeisung).DruckFest = False: kc(einspeisung).DruckBarA = 0
        kc(einspeisung).FlussVorgegeben = True: kc(einspeisung).ExternQ = q
        For i = 1 To UBound(kc)
            If kc(i).DruckFest Then fest = fest + 1
            If i <> einspeisung And kc(i).ExternQ > 0 Then NetzFehler "Weitere feste Einspeisestroeme: keine Gesamt-Systemkennlinie. Alternative: Netzwerkbetriebspunkt."
        Next i
        If fest = 0 Then NetzFehler "Nach Q-Vorgabe fehlt die Druckreferenz. Auslassdruck festlegen oder Netzwerkbetriebspunkt berechnen."
        NetzLoesen kc, e, rho, mu, pq, ph, aus, maxIter, tq, tp
        If aus.Erfolg Then druck = aus.Druck(einspeisung)
    End If
End Sub

Public Sub NetzwerkKennlinieBerechnen()
    On Error GoTo Fehler
    Dim k() As NetzKnoten, e() As NetzLeitung, pq() As Double, ph() As Double, r As NetzResultat
    Dim rho As Double, mu As Double, qmin As Double, qmax As Double, nval As Double, q As Double, dp As Double
    Dim ws As Worksheet, cfg As Worksheet, n As Long, i As Long, j As Long, feed As Long, pump As Long
    Dim lim As Long, tq As Double, tp As Double, data() As Variant, flows() As Variant, id As String, message As String
    Dim ch As ChartObject, s As Series
    Set ws = ThisWorkbook.Worksheets("Netzwerk_Kennlinie"): Set cfg = ThisWorkbook.Worksheets("Netzwerk_Setup")
    NetzAusgabeLeeren ws
    ws.Range("A2").Value = "NICHT AKTUELL: Kennlinie laeuft"
    NetzwerkEinlesen k, e, rho, mu, pq, ph
    NetzToleranzen lim, tq, tp
    id = CStr(cfg.Range("B17").Value)
    For i = 1 To UBound(k)
        If StrComp(id, k(i).ID, vbTextCompare) = 0 Then feed = i
    Next i
    qmin = Zahl(cfg.Range("B18"), "Kennlinie Qmin"): qmax = Zahl(cfg.Range("B19"), "Kennlinie Qmax")
    nval = Zahl(cfg.Range("B20"), "Kennlinienpunkte")
    If qmin < 0 Or qmax <= qmin Or nval < 2 Or nval > 1000 Or nval <> Fix(nval) Then NetzFehler "Kennlinie: 0<=Qmin<Qmax, 2..1000 ganzzahlige Punkte."
    n = CLng(nval): ReDim data(1 To n, 1 To 5): ReDim flows(1 To n, 1 To UBound(e) + 1)
    For i = 1 To UBound(e)
        If e(i).Aktiv And e(i).Pumpe Then pump = i
    Next i
    If pump > 0 Then
        PruefePumpendaten pq, ph, rho
        If pq(LBound(pq)) <> 0 Then NetzFehler "Netzpumpe muss bei Q=0 beginnen."
    End If
    For i = 1 To n
        q = qmin + (qmax - qmin) * (i - 1) / (n - 1)
        NetzKennlinienPunkt k, e, feed, q, rho, mu, pq, ph, r, dp, lim, tq, tp
        If Not r.Erfolg Then NetzFehler "Q=" & q & ": " & r.Status & ". Alternative: Netzwerkbetriebspunkt berechnen."
        data(i, 1) = q: data(i, 2) = dp: data(i, 3) = dp * 100000# / (rho * G)
        data(i, 4) = r.MaxBilanz: data(i, 5) = r.MaxEnergie
        flows(i, 1) = q
        For j = 1 To UBound(e)
            flows(i, j + 1) = r.Q(j)
        Next j
    Next i
    ws.Range("A4:E4").Value = Array("Q Port/Pumpe m3/h", "Druck bar", "Druckhoehe m", "Bilanz m3/h", "Energiefehler bar")
    ws.Range("A5").Resize(n, 5).Value = data
    ws.Range("AA4").Value = "Q gesamt/Port"
    For j = 1 To UBound(e)
        ws.Cells(4, 27 + j).Value = e(j).ID
    Next j
    ws.Range("AA5").Resize(n, UBound(e) + 1).Value = flows
    ws.Range("A3").Value = "Ohne Pumpe: p am Port bar(a), Druckhoehe absolut. Mit Pumpe: erforderliche Pumpen-dp/H. Verteilung pro Punkt rechts dokumentiert."
    Set ch = ws.ChartObjects.Add(ws.Range("G5").Left, ws.Range("G5").Top, 700, 400): ch.Name = "net_curve"
    With ch.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        Set s = .SeriesCollection.NewSeries
        s.Name = "Erforderlicher Systemdruck": s.XValues = ws.Range("A5").Resize(n, 1): s.Values = ws.Range("B5").Resize(n, 1)
        If pump > 0 Then
            ws.Range("H30:I30").Value = Array("Pumpen Q", "Pumpendruck bar")
            For i = LBound(pq) To UBound(pq)
                ws.Cells(30 + i, 8).Value = pq(i): ws.Cells(30 + i, 9).Value = rho * G * ph(i) / 100000#
            Next i
            Set s = .SeriesCollection.NewSeries
            s.Name = "Herstellerpumpe": s.XValues = ws.Range("H31").Resize(UBound(pq), 1): s.Values = ws.Range("I31").Resize(UBound(pq), 1)
            NetzLoesen k, e, rho, mu, pq, ph, r, lim, tq, tp
            If r.Erfolg Then
                ws.Range("K31").Value = r.PumpenQ: ws.Range("L31").Value = r.PumpenDruck
                Set s = .SeriesCollection.NewSeries
                s.Name = "Gekoppelter Betriebspunkt": s.XValues = ws.Range("K31"): s.Values = ws.Range("L31")
                s.ChartType = xlXYScatter: s.MarkerStyle = xlMarkerStyleCircle: s.MarkerSize = 9
            Else
                ws.Range("G3").Value = "Kein Hersteller-Betriebspunkt: " & r.Status
            End If
        End If
        .HasTitle = True: .ChartTitle.Text = "Netzwerk-Systemkennlinie"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Q [m3/h]"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "Druck [bar]"
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
    End With
    ws.Range("A2").Value = "BERECHNET: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "; Einspeiseport " & id
    ws.Activate
    Exit Sub
Fehler:
    message = Err.Description
    On Error Resume Next
    NetzAusgabeLeeren ws
    ws.Range("A2").Value = "UNGUELTIG: " & message
    MsgBox message, vbExclamation, "Netzwerkkennlinie"
End Sub
