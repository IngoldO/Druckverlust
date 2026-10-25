Attribute VB_Name = "NetzBericht"
Option Explicit

Public Sub NetzKennzahlen(ByVal ws As Worksheet, ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByRef r As NetzResultat, ByVal rho As Double, ByVal mu As Double)
    Dim i As Long, zufluss As Double, pmin As Double, pmax As Double, maxloss As Double, loss As Double
    pmin = r.Druck(1): pmax = pmin
    For i = 1 To UBound(k)
        If r.Druck(i) < pmin Then pmin = r.Druck(i)
        If r.Druck(i) > pmax Then pmax = r.Druck(i)
        If k(i).DruckFest Then
            If r.Bilanz(i) < 0 Then zufluss = zufluss - r.Bilanz(i)
        ElseIf k(i).ExternQ > 0 Then
            zufluss = zufluss + k(i).ExternQ
        End If
    Next i
    For i = 1 To UBound(e)
        If e(i).Aktiv Then
            loss = Abs(NetzVerlust(r.Q(i), e(i), rho, mu))
            If loss > maxloss Then maxloss = loss
        End If
    Next i
    ws.Range("P6:Q6").Value = Array("Kennzahl", "Wert")
    ws.Range("P7:Q7").Value = Array("Externer Gesamtzufluss m3/h", zufluss)
    ws.Range("P8:Q8").Value = Array("Minimaler Knotendruck bar(a)", pmin)
    ws.Range("P9:Q9").Value = Array("Maximaler Knotendruck bar(a)", pmax)
    ws.Range("P10:Q10").Value = Array("Groesster irreversibler Verlust bar", maxloss)
    ws.Range("P11:Q11").Value = Array("Pumpen-Q m3/h (0 ohne Pumpe)", r.PumpenQ)
    ws.Range("P12:Q12").Value = Array("Pumpen-dp bar", r.PumpenDruck)
    ws.Range("P13:Q13").Value = Array("Pumpenfoerderhoehe m", r.PumpenDruck * 100000# / (rho * G))
    ws.Range("P14:Q14").Value = Array("Freies Massenresiduum kg/s", r.MaxMasse)
    ws.Range("P15:Q15").Value = Array("Solverstatus", r.Status)
End Sub

Private Sub BerichtKoerperLeeren(ByVal ws As Worksheet)
    Dim last As Long
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If last < 6 Then last = 6
    ws.Range("A6:Z" & last).UnMerge
    ws.Range("A6:Z" & last).ClearContents
End Sub

Private Sub BerichtZeile(ByVal ws As Worksheet, ByRef row As Long, ByVal text As String)
    ws.Cells(row, 1).Resize(1, 12).Merge
    With ws.Cells(row, 1)
        .Value = text: .WrapText = True: .Font.Size = 10
        .Font.Bold = False: .Font.Color = vbBlack: .Interior.Color = vbWhite
        If Left(text, 8) = "KNOTEN |" Or Left(text, 11) = "LEITUNGEN |" Or Left(text, 11) = "ARMATUREN |" Or Left(text, 19) = "PUMPENSTUETZPUNKTE |" Or Left(text, 19) = "NETZWERK-KENNLINIE |" Then
            .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(24, 84, 105)
        End If
    End With
    ws.Rows(row).RowHeight = 14 * (2 + Len(text) \ 110)
    row = row + 1
End Sub

Public Sub NetzwerkBerichtErstellen()
    On Error GoTo Fehler
    Dim ws As Worksheet, k() As NetzKnoten, e() As NetzLeitung, pq() As Double, ph() As Double
    Dim rho As Double, mu As Double, lim As Long, tq As Double, tp As Double, r As NetzResultat
    Dim row As Long, i As Long, j As Long, pump As Boolean, status As String, armaturenText As String
    Dim t As ListObject, lr As ListRow, p As Betriebspunkt, ch As ChartObject, s As Series, richtung As String
    Dim aktiv As Boolean, wirksam As Double, zeilenhoehe As Double, tempHoehe As Double
    Dim data() As Variant, kc() As NetzKnoten, cr As NetzResultat, dp As Double, q As Double
    Dim qmin As Double, qmax As Double, n As Long, nval As Double, feed As Long, id As String
    Set ws = ThisWorkbook.Worksheets("Netzwerk_Bericht")
    ' Keep document metadata and fixed header; rebuild only the generated body.
    NetzShapesLoeschen ws
    BerichtKoerperLeeren ws
    ws.Range("A6").Value = "NICHT AKTUELL: Bericht wird neu berechnet"
    NetzwerkEinlesen k, e, rho, mu, pq, ph: NetzToleranzen lim, tq, tp
    NetzLoesen k, e, rho, mu, pq, ph, r, lim, tq, tp
    If Not r.Erfolg Then NetzFehler r.Status
    ws.Range("A7:L7").Merge
    ws.Range("A7").Value = "rho=" & rho & " kg/m3; mu=" & mu * 1000 & " mPa s; Iterationen=" & r.Iterationen & "; Bilanz=" & r.MaxBilanz & " m3/h; Energie=" & r.MaxEnergie & " bar"
    ws.Range("A7").WrapText = True: ws.Rows(7).RowHeight = 30
    NetzSchemaZeichnen ws, k, e, r, ws.Range("A9").Top, ws.Range("A1:L1").Width - 20, 240
    row = 26
    BerichtZeile ws, row, "NETZWERKSCHEMA: schematisch, nicht massstaeblich. Fuer grosse Netze separates Blatt Netzwerk_Diagramm vergroessert drucken."
    NetzKennzahlen ws, k, e, r, rho, mu
    For i = 7 To 15
        BerichtZeile ws, row, CStr(ws.Cells(i, 16).Value) & ": " & CStr(ws.Cells(i, 17).Value)
    Next i
    BerichtZeile ws, row, "Solver: " & r.Status & "; freies Massenresiduum=" & r.MaxMasse & " kg/s; Pivotdiagnose=" & r.PivotVerhaeltnis
    BerichtZeile ws, row, "FORMELN: sum(Qein)-sum(Qaus)+Qextern=0; Massenbilanz=rho*Volumenbilanz/3600."
    BerichtZeile ws, row, "p_i-p_j+rho*g*(z_i-z_j)/100000+dpPumpe-sgn(Q)*(dpRohr+dpZeta+dpKv+dpBlende)=0 [bar]."
    BerichtZeile ws, row, "Darcy: dpR=lambda*(L/D)*rho*v^2/2; Re=rho*v*D/mu; laminar lambda=64/Re; turbulent Colebrook."
    BerichtZeile ws, row, "Kv: dp[bar]=n*(rho/1000)*(Q/Kv)^2. Blende: zeta=(sqrt(1-beta^4*(1-Cd^2))/(Cd*beta^2)-1)^2."
    BerichtZeile ws, row, "ANNAHMEN: stationaer, einphasig, inkompressibel, Newtonsch, konstante Stoffwerte. Re 2300-4000 unsicher. Zeta/Kv/Cd projektbezogen pruefen; keine Kavitation-, NPSH-, Thermo- oder Festigkeitspruefung."
    BerichtZeile ws, row, "Druckknoten: Bilanz ist Leitungs-Nettozufluss; externer Reservoirzufluss=-Bilanz. Geschlossene Rueckschlagventile koennen Druck halten; Residuum misst die Verletzung der Sperrbedingung."
    BerichtZeile ws, row, "Geprueft durch: __________________  Datum: ______________  Freigabe: __________________"
    ws.ResetAllPageBreaks: ws.HPageBreaks.Add Before:=ws.Cells(row + 1, 1)
    row = row + 1
    BerichtZeile ws, row, "KNOTEN | RAND- UND ERGEBNISDATEN (bar(a), m, m3/h)"
    ws.Cells(row, 1).Resize(1, 8).Value = Array("ID", "Bezeichnung", "Typ", "z m", "p bar(a)", "p fest", "Extern Q", "Bilanz Q")
    row = row + 1
    For i = 1 To UBound(k)
        ws.Cells(row, 1).Resize(1, 8).Value = Array(k(i).ID, k(i).Bezeichnung, k(i).Typ, k(i).Hoehe, r.Druck(i), k(i).DruckFest, k(i).ExternQ, r.Bilanz(i))
        zeilenhoehe = 32
        For j = 1 To 8
            tempHoehe = 14 * (1 + Len(CStr(ws.Cells(row, j).Value)) / 9)
            If tempHoehe > zeilenhoehe Then zeilenhoehe = tempHoehe
        Next j
        If zeilenhoehe > 409 Then zeilenhoehe = 409
        ws.Rows(row).RowHeight = zeilenhoehe
        With ws.Cells(row, 1).Resize(1, 12)
            .WrapText = True: .Font.Size = 10: .Font.Bold = False: .Font.Color = vbBlack: .Interior.Color = vbWhite
        End With
        row = row + 1
    Next i
    BerichtZeile ws, row, "LEITUNGEN | Eingaben und Resultate; alle Druckverluste als positive Betraege"
    For i = 1 To UBound(e)
        If e(i).Aktiv Then
            p = BerechnePunkt(Abs(r.Q(i)), e(i).D, e(i).L, e(i).Eps, rho, mu, 0, e(i).Zeta, 0)
            BerichtZeile ws, row, e(i).ID & " | " & e(i).Bezeichnung & ": " & k(e(i).Von).ID & " -> " & k(e(i).Nach).ID & "; DN=" & e(i).DN & "; Di=" & e(i).D * 1000 & " mm; L=" & e(i).L & " m; eps=" & e(i).Eps * 1000 & " mm"
            richtung = "Start -> Ende": If r.Q(i) < 0 Then richtung = "Ende -> Start"
            If Abs(r.Q(i)) <= tq Then richtung = "Nahe Null"
            BerichtZeile ws, row, "Richtung=" & richtung & "; Q=" & r.Q(i) & " m3/h; v=" & Sgn(r.Q(i)) * p.V & " m/s; Re=" & p.Re & "; lambda=" & p.Lambda & "; Rohr=" & p.RohrBar & " bar; Zeta=" & p.EinzelBar & " bar"
            BerichtZeile ws, row, "Kv=" & e(i).ZKv * rho * p.V ^ 2 / 200000# & " bar; Blende=" & e(i).ZBlende * rho * p.V ^ 2 / 200000# & " bar; irreversibel=" & Abs(NetzVerlust(r.Q(i), e(i), rho, mu)) & " bar; Energiefehler=" & r.Energie(i) & " bar"
            BerichtZeile ws, row, "Wirksame Beiwerte am Leitungs-D: Zeta=" & e(i).Zeta & "; ZKv=" & e(i).ZKv & "; ZBlende=" & e(i).ZBlende
            BerichtZeile ws, row, "Warnungen: " & p.Hinweis & " Konstante Zeta/Kv/Cd, lokale Blenden-Re und Reduzierverluste separat pruefen."
        Else
            BerichtZeile ws, row, e(i).ID & ": INAKTIV"
        End If
        If e(i).Aktiv And e(i).Pumpe Then pump = True
    Next i
    BerichtZeile ws, row, "ARMATUREN | vollstaendige Zuordnung und Eingaben; aktive Modelle werden genau einmal addiert"
    Set t = ThisWorkbook.Worksheets("Netzwerk_Armaturen").ListObjects("tblNetzArmaturen")
    For Each lr In t.ListRows
        armaturenText = ""
        For j = 1 To t.ListColumns.Count
            armaturenText = armaturenText & t.ListColumns(j).Name & "=" & CStr(lr.Range.Cells(1, j).Value) & "; "
        Next j
        aktiv = True
        If Not IsEmpty(lr.Range.Cells(1, 10).Value) Then aktiv = (Zahl(lr.Range.Cells(1, 10), "Aktiv") = 1)
        If aktiv Then
            If Zahl(lr.Range.Cells(1, 4), "Anzahl") > 0 Then
                If CStr(lr.Range.Cells(1, 2).Value) = "Zeta" Or CStr(lr.Range.Cells(1, 2).Value) = "Rueckschlag" Then
                    If IsEmpty(lr.Range.Cells(1, 5).Value) Then wirksam = StandardZeta(CStr(lr.Range.Cells(1, 3).Value)) Else wirksam = Zahl(lr.Range.Cells(1, 5), "Zeta")
                    armaturenText = armaturenText & "wirksames Zeta je Armatur=" & wirksam
                End If
            End If
        End If
        BerichtZeile ws, row, armaturenText
    Next lr
    If pump Then
        BerichtZeile ws, row, "PUMPENSTUETZPUNKTE | konstante Drehzahl; linear zwischen Punkten, keine Extrapolation"
        For i = LBound(pq) To UBound(pq)
            BerichtZeile ws, row, "Q=" & pq(i) & " m3/h; H=" & ph(i) & " m"
        Next i
    End If
    ' Curve is computed from the SAME current data. A failed curve does not falsify a valid operating report.
    id = CStr(ThisWorkbook.Worksheets("Netzwerk_Setup").Range("B17").Value)
    For i = 1 To UBound(k)
        If StrComp(id, k(i).ID, vbTextCompare) = 0 Then feed = i
    Next i
    status = BerichtKennlinie(ws, row, k, e, rho, mu, pq, ph, r, feed, lim, tq, tp)
    BerichtZeile ws, row, "Kennlinienstatus: " & status
    ws.Range("A6:L6").Merge
    ws.Range("A6").Value = "BERECHNET: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "; Eingabesnapshot; nach Aenderungen neu erstellen."
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(row, 12)).Address
    ws.PageSetup.Zoom = 85
    ws.Activate
    Exit Sub
Fehler:
    status = Err.Description
    On Error Resume Next
    NetzShapesLoeschen ws
    BerichtKoerperLeeren ws
    ws.Range("A6").Value = "UNGUELTIG: " & status
    MsgBox status, vbExclamation, "Netzwerkbericht"
End Sub

Private Function BerichtKennlinie(ByVal ws As Worksheet, ByRef row As Long, ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByVal rho As Double, ByVal mu As Double, ByRef pq() As Double, ByRef ph() As Double, ByRef op As NetzResultat, ByVal feed As Long, ByVal lim As Long, ByVal tq As Double, ByVal tp As Double) As String
    On Error GoTo KeineKennlinie
    Dim cfg As Worksheet, nval As Double, n As Long, qmin As Double, qmax As Double, q As Double, dp As Double
    Dim cr As NetzResultat, data() As Variant, i As Long, j As Long, pump As Boolean, ch As ChartObject, s As Series
    Set cfg = ThisWorkbook.Worksheets("Netzwerk_Setup")
    qmin = Zahl(cfg.Range("B18"), "Qmin"): qmax = Zahl(cfg.Range("B19"), "Qmax"): nval = Zahl(cfg.Range("B20"), "Punkte")
    If qmin < 0 Or qmax <= qmin Or nval < 2 Or nval > 1000 Or nval <> Fix(nval) Then NetzFehler "Kennlinienbereich ungueltig."
    n = CLng(nval): ReDim data(1 To n, 1 To 2)
    For i = 1 To n
        q = qmin + (qmax - qmin) * (i - 1) / (n - 1)
        NetzKennlinienPunkt k, e, feed, q, rho, mu, pq, ph, cr, dp, lim, tq, tp
        If Not cr.Erfolg Then NetzFehler cr.Status
        data(i, 1) = q: data(i, 2) = dp
    Next i
    ws.Range("N30").Resize(n, 2).Value = data
    For i = 1 To UBound(e)
        If e(i).Aktiv And e(i).Pumpe Then pump = True
    Next i
    ws.HPageBreaks.Add Before:=ws.Cells(row + 1, 1): row = row + 1
    BerichtZeile ws, row, "NETZWERK-KENNLINIE | Einspeiseport " & k(feed).ID & "; Q=" & qmin & ".." & qmax & "; Punkte=" & n
    Set ch = ws.ChartObjects.Add(ws.Cells(row, 1).Left, ws.Cells(row, 1).Top, ws.Range("A1:L1").Width - 10, 340): ch.Name = "net_report_curve"
    With ch.Chart
        .ChartType = xlXYScatterLinesNoMarkers: .PlotVisibleOnly = False
        Set s = .SeriesCollection.NewSeries
        s.Name = "System": s.XValues = ws.Range("N30").Resize(n, 1): s.Values = ws.Range("O30").Resize(n, 1)
        If pump Then
            For i = LBound(pq) To UBound(pq)
                ws.Cells(30 + i, 18).Value = pq(i): ws.Cells(30 + i, 19).Value = rho * G * ph(i) / 100000#
            Next i
            Set s = .SeriesCollection.NewSeries
            s.Name = "Pumpe": s.XValues = ws.Range("R31").Resize(UBound(pq), 1): s.Values = ws.Range("S31").Resize(UBound(pq), 1)
            ws.Range("T31").Value = op.PumpenQ: ws.Range("U31").Value = op.PumpenDruck
            Set s = .SeriesCollection.NewSeries
            s.Name = "Betriebspunkt": s.XValues = ws.Range("T31"): s.Values = ws.Range("U31")
            s.ChartType = xlXYScatter: s.MarkerStyle = xlMarkerStyleCircle: s.MarkerSize = 8
        End If
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Q [m3/h]"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "Druck [bar]"
        .HasLegend = True: .Legend.Position = xlLegendPositionBottom
    End With
    row = row + 24
    BerichtZeile ws, row, "Ohne Pumpe: absoluter Einspeisedruck. Mit Pumpe: erforderliche Druckerhoehung am Pumpenelement. Andere Randbedingungen unveraendert; keine Hersteller-Extrapolation."
    BerichtKennlinie = "berechnet aus demselben aktuellen Netzstand"
    Exit Function
KeineKennlinie:
    Dim fehlertext As String
    fehlertext = Err.Description
    On Error Resume Next
    ws.ChartObjects("net_report_curve").Delete
    BerichtKennlinie = "NICHT VERFUEGBAR: " & fehlertext & ". Alternative: dokumentierter Netzwerkbetriebspunkt."
End Function
