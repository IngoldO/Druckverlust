Attribute VB_Name = "NetzExcel"
Option Explicit

Private Function IndexID(ByVal id As String, ByRef ids() As String) As Long
    Dim i As Long
    For i = 1 To UBound(ids)
        If StrComp(id, ids(i), vbTextCompare) = 0 Then IndexID = i: Exit Function
    Next i
    NetzFehler "Unbekannte ID: " & id
End Function

Private Function Schalter(ByVal c As Range, ByVal txt As String, Optional ByVal standard As Boolean = True) As Boolean
    Dim x As Double
    If IsEmpty(c.Value) Then Schalter = standard: Exit Function
    x = Zahl(c, txt)
    If x <> 0 And x <> 1 Then NetzFehler txt & ": nur 0 oder 1."
    Schalter = (x = 1)
End Function

Public Sub NetzwerkEinlesen(ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByRef rho As Double, ByRef mu As Double, ByRef pq() As Double, ByRef ph() As Double)
    Dim t As ListObject, lr As ListRow, r As Range, i As Long, j As Long, count As Long
    Dim ids() As String, leitungen() As String, n As Double, modell As String, z As Variant
    Dim kv As Double, lokal As Double, loch As Double, cd As Double, typ As String, dn As Double
    With ThisWorkbook.Worksheets("Netzwerk_Setup")
        rho = Zahl(.Range("B5"), "Netz-Dichte")
        mu = Zahl(.Range("B6"), "Netz-Viskositaet") / 1000
    End With
    Set t = ThisWorkbook.Worksheets("Netzwerk_Knoten").ListObjects("tblNetzKnoten")
    count = t.ListRows.Count
    If count < 2 Then NetzFehler "Mindestens zwei Knoten erforderlich."
    ReDim k(1 To count): ReDim ids(1 To count)
    For Each lr In t.ListRows
        i = i + 1: Set r = lr.Range
        k(i).ID = Trim(CStr(r.Cells(1, 1).Value)): ids(i) = k(i).ID
        k(i).Bezeichnung = CStr(r.Cells(1, 2).Value): k(i).Typ = CStr(r.Cells(1, 3).Value)
        Select Case k(i).Typ
        Case "Einspeisung", "Verteiler", "Verbindungsknoten", "Sammelkollektor", "Verbraucher", "Behaelter", "Offener Auslass"
        Case Else: NetzFehler "Unbekannter Knotentyp: " & k(i).Typ
        End Select
        k(i).Hoehe = Zahl(r.Cells(1, 4), "Knotenhoehe")
        k(i).DruckFest = Not IsEmpty(r.Cells(1, 5).Value)
        If k(i).DruckFest Then k(i).DruckBarA = Zahl(r.Cells(1, 5), "Absolutdruck")
        k(i).FlussVorgegeben = Not IsEmpty(r.Cells(1, 6).Value)
        If k(i).FlussVorgegeben Then k(i).ExternQ = Zahl(r.Cells(1, 6), "Externer Fluss")
        If k(i).Typ = "Behaelter" Or k(i).Typ = "Offener Auslass" Then
            If Not k(i).DruckFest Then NetzFehler "Behaelter/offener Auslass benoetigt expliziten Absolutdruck."
        End If
        If IsEmpty(r.Cells(1, 7).Value) Xor IsEmpty(r.Cells(1, 8).Value) Then NetzFehler "Manuelle Position: X und Y gemeinsam eingeben."
        k(i).PositionManuell = Not IsEmpty(r.Cells(1, 7).Value)
        If Not IsEmpty(r.Cells(1, 7).Value) Then k(i).X = Zahl(r.Cells(1, 7), "Position X")
        If Not IsEmpty(r.Cells(1, 8).Value) Then k(i).Y = Zahl(r.Cells(1, 8), "Position Y")
    Next lr
    Set t = ThisWorkbook.Worksheets("Netzwerk_Leitungen").ListObjects("tblNetzLeitungen")
    count = t.ListRows.Count
    If count < 1 Then NetzFehler "Mindestens eine Leitung erforderlich."
    ReDim e(1 To count): ReDim leitungen(1 To count): i = 0
    For Each lr In t.ListRows
        i = i + 1: Set r = lr.Range
        e(i).ID = Trim(CStr(r.Cells(1, 1).Value)): leitungen(i) = e(i).ID
        e(i).Bezeichnung = CStr(r.Cells(1, 2).Value)
        e(i).Aktiv = Schalter(r.Cells(1, 9), "Leitung aktiv")
        If e(i).Aktiv Then
            e(i).Von = IndexID(CStr(r.Cells(1, 3).Value), ids): e(i).Nach = IndexID(CStr(r.Cells(1, 4).Value), ids)
            dn = Zahl(r.Cells(1, 5), "DN")
            If Application.CountIf(ThisWorkbook.Names("DNWerte").RefersToRange, dn) <> 1 Then NetzFehler "DN nicht in Auswahltabelle."
            e(i).DN = dn: e(i).D = Zahl(r.Cells(1, 6), "Di") / 1000
            e(i).L = Zahl(r.Cells(1, 7), "Laenge"): e(i).Eps = Zahl(r.Cells(1, 8), "Rauigkeit") / 1000
            e(i).Pumpe = Schalter(r.Cells(1, 10), "Leitung Pumpe", False)
        End If
    Next lr
    Set t = ThisWorkbook.Worksheets("Netzwerk_Armaturen").ListObjects("tblNetzArmaturen")
    For Each lr In t.ListRows
        Set r = lr.Range
        If Schalter(r.Cells(1, 10), "Armatur aktiv") Then
            j = IndexID(CStr(r.Cells(1, 1).Value), leitungen)
            If Not e(j).Aktiv Then NetzFehler "Aktive Armatur auf deaktivierter Leitung."
            modell = CStr(r.Cells(1, 2).Value): typ = CStr(r.Cells(1, 3).Value)
            n = Zahl(r.Cells(1, 4), "Anzahl")
            If n = 0 Then GoTo NextArmatur
            z = Empty: kv = 0: lokal = 0: loch = 0: cd = 0
            If Not IsEmpty(r.Cells(1, 5).Value) Then z = Zahl(r.Cells(1, 5), "Zeta manuell")
            If modell = "Kv" Then kv = Zahl(r.Cells(1, 6), "Kv")
            If modell = "Blende" Then
                lokal = Zahl(r.Cells(1, 7), "Blenden-Di") / 1000
                loch = Zahl(r.Cells(1, 8), "Loch-D") / 1000
                cd = Zahl(r.Cells(1, 9), "Cd")
            End If
            NetzArmaturAddieren e(j), modell, typ, n, rho, z, kv, lokal, loch, cd
        End If
NextArmatur:
    Next lr
    count = 0
    For i = 1 To UBound(e)
        If e(i).Aktiv And e(i).Pumpe Then count = count + 1
    Next i
    ReDim pq(1 To 1): ReDim ph(1 To 1)
    If count > 0 Then
        Set t = ThisWorkbook.Worksheets("Netzwerk_Pumpe").ListObjects("tblNetzPumpe")
        count = t.ListRows.Count
        If count < 2 Or count > 1000 Then NetzFehler "Netzpumpe: 2 bis 1000 Punkte."
        ReDim pq(1 To count): ReDim ph(1 To count): i = 0
        For Each lr In t.ListRows
            i = i + 1
            pq(i) = Zahl(lr.Range.Cells(1, 1), "Pumpen-Q"): ph(i) = Zahl(lr.Range.Cells(1, 2), "Pumpen-H")
        Next lr
    End If
End Sub

Public Sub NetzwerkBerechnen()
    On Error GoTo Fehler
    Dim k() As NetzKnoten, e() As NetzLeitung, pq() As Double, ph() As Double, r As NetzResultat
    Dim rho As Double, mu As Double, lim As Double, tq As Double, tp As Double, ws As Worksheet, i As Long, row As Long
    Dim p As Betriebspunkt, data(1 To 1, 1 To 14) As Variant, msg As String
    Set ws = ThisWorkbook.Worksheets("Netzwerk_Ergebnisse")
    ws.Cells.ClearContents
    ws.Range("A2").Value = "NICHT AKTUELL: Berechnung laeuft"
    NetzwerkEinlesen k, e, rho, mu, pq, ph
    With ThisWorkbook.Worksheets("Netzwerk_Setup")
        lim = Zahl(.Range("B7"), "Iterationslimit")
        tq = Zahl(.Range("B8"), "Volumentoleranz"): tp = Zahl(.Range("B9"), "Drucktoleranz")
    End With
    If lim < 1 Or lim > 1000 Or lim <> Fix(lim) Then NetzFehler "Iterationslimit ganzzahlig 1..1000."
    NetzLoesen k, e, rho, mu, pq, ph, r, CLng(lim), tq, tp
    If Not r.Erfolg Then NetzFehler r.Status
    ws.Range("A2").Value = r.Status & "; Iterationen=" & r.Iterationen
    ws.Range("A3").Value = "Max. freie Knotenbilanz=" & r.MaxBilanz & " m3/h; " & r.MaxMasse & " kg/s; Energie=" & r.MaxEnergie & " bar"
    ws.Range("A4").Value = "Pivotverhaeltnis=" & r.PivotVerhaeltnis & "; Druckknoten-Bilanzen sind Reservoirreaktionen, keine Solverfehler."
    ws.Range("A6:F6").Value = Array("Knoten", "Bezeichnung", "p bar(a)", "Hoehe m", "Bilanz m3/h", "Druck fest")
    For i = 1 To UBound(k)
        ws.Cells(6 + i, 1).Resize(1, 6).Value = Array(k(i).ID, k(i).Bezeichnung, r.Druck(i), k(i).Hoehe, r.Bilanz(i), k(i).DruckFest)
    Next i
    row = 9 + UBound(k)
    ws.Cells(row, 1).Resize(1, 14).Value = Array("Leitung", "Start", "Ende", "Q m3/h", "v m/s signiert", "Re", "Darcy Lambda", "Rohr bar", "Zeta bar", "Kv bar", "Blende bar", "Irreversibel bar", "Energie-Residuum bar", "Status / Richtung")
    For i = 1 To UBound(e)
        If e(i).Aktiv Then
            p = BerechnePunkt(Abs(r.Q(i)), e(i).D, e(i).L, e(i).Eps, rho, mu, 0, e(i).Zeta, 0)
            data(1, 1) = e(i).ID: data(1, 2) = k(e(i).Von).ID: data(1, 3) = k(e(i).Nach).ID
            data(1, 4) = r.Q(i): data(1, 5) = Sgn(r.Q(i)) * p.V: data(1, 6) = p.Re: data(1, 7) = p.Lambda
            data(1, 8) = p.RohrBar: data(1, 9) = p.EinzelBar
            data(1, 10) = e(i).ZKv * rho * p.V ^ 2 / 200000#: data(1, 11) = e(i).ZBlende * rho * p.V ^ 2 / 200000#
            data(1, 12) = Abs(NetzVerlust(r.Q(i), e(i), rho, mu)): data(1, 13) = r.Energie(i)
            msg = "Start -> Ende": If r.Q(i) < 0 Then msg = "Ende -> Start"
            If Abs(r.Q(i)) <= tq Then msg = "Nahe Null"
            data(1, 14) = msg & "; " & p.Hinweis & " Kv/Cd und lokale Durchmesser projektbezogen pruefen."
            ws.Cells(row + i, 1).Resize(1, 14).Value = data
        Else
            ws.Cells(row + i, 1).Value = e(i).ID: ws.Cells(row + i, 14).Value = "INAKTIV"
        End If
    Next i
    NetzKennzahlen ws, k, e, r, rho, mu
    With ThisWorkbook.Worksheets("Netzwerk_Diagramm")
        NetzSchemaZeichnen ThisWorkbook.Worksheets("Netzwerk_Diagramm"), k, e, r
        .Range("A2").Value = "BERECHNET: " & Format(Now, "yyyy-mm-dd hh:nn:ss") & "; Pfeile zeigen tatsaechliche Stroemungsrichtung."
    End With
    ws.Range("A2").Value = ws.Range("A2").Value & "; " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Activate
    Exit Sub
Fehler:
    msg = Err.Description
    On Error Resume Next
    ws.Cells.ClearContents
    ws.Range("A2").Value = "UNGUELTIG: " & msg
    NetzShapesLoeschen ThisWorkbook.Worksheets("Netzwerk_Diagramm")
    ThisWorkbook.Worksheets("Netzwerk_Diagramm").Range("A2").Value = "UNGUELTIG: " & msg
    MsgBox msg, vbExclamation, "Netzwerk"
End Sub
