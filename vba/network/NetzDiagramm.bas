Attribute VB_Name = "NetzDiagramm"
Option Explicit

Public Sub NetzSchemaZeichnen(ByVal ws As Worksheet, ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByRef r As NetzResultat, Optional ByVal oben As Double = 75, Optional ByVal maxBreite As Double = 0, Optional ByVal maxHoehe As Double = 0)
    Dim layer() As Long, xpos() As Double, ypos() As Double
    Dim i As Long, j As Long, a As Long, b As Long, n As Long
    Dim schemaBreite As Double, schemaHoehe As Double, layoutFaktor As Double, sx As Double
    Dim rechteGrenze As Double, untereGrenze As Double, fontsize As Double
    Dim sy As Double, duplicate As Long, node As Shape, verbindung As Shape, beschriftung As Shape, pump As Shape
    If Not r.Erfolg Then NetzFehler "Kein gueltiger Betriebspunkt fuer das Schema."
    NetzShapesLoeschen ws
    n = UBound(k): ReDim layer(1 To n): ReDim xpos(1 To n): ReDim ypos(1 To n)
    NetzLayoutBerechnen k, e, layer, xpos, ypos
    For i = 1 To n
        If xpos(i) + 140 > schemaBreite Then schemaBreite = xpos(i) + 140
        If ypos(i) + 65 > schemaHoehe Then schemaHoehe = ypos(i) + 65
    Next i
    ' Include labels/pump symbols in the fitting bounds, also for parallel edges.
    For i = 1 To UBound(e)
        If e(i).Aktiv Then
            a = e(i).Von: b = e(i).Nach: duplicate = 0
            For j = 1 To i - 1
                If (e(j).Von = a And e(j).Nach = b) Or (e(j).Von = b And e(j).Nach = a) Then duplicate = duplicate + 1
            Next j
            sx = (xpos(a) + xpos(b) + 140) / 2 + 75
            sy = (ypos(a) + ypos(b)) / 2 + 24 * duplicate + 65
            If sx > schemaBreite Then schemaBreite = sx
            If sy > schemaHoehe Then schemaHoehe = sy
        End If
    Next i
    layoutFaktor = 1
    If maxBreite > 0 And schemaBreite > maxBreite Then layoutFaktor = maxBreite / schemaBreite
    If maxHoehe > 0 And schemaHoehe * layoutFaktor > maxHoehe Then layoutFaktor = maxHoehe / schemaHoehe
    For i = 1 To n
        sx = 10 + xpos(i) * layoutFaktor: sy = oben + ypos(i) * layoutFaktor
        Set node = ws.Shapes.AddShape(msoShapeRoundedRectangle, sx, sy, 140 * layoutFaktor, 65 * layoutFaktor)
        node.Name = "net_node_" & i
        node.TextFrame.Characters.Text = k(i).ID & " | " & k(i).Bezeichnung & vbLf & k(i).Typ & vbLf & Format(r.Druck(i), "0.0000") & " bar(a)"
        fontsize = 9 * layoutFaktor: If fontsize < 1 Then fontsize = 1
        node.TextFrame.Characters.Font.Size = fontsize
        node.TextFrame.Characters.Font.Color = RGB(20, 45, 55)
        node.Fill.ForeColor.RGB = RGB(225, 241, 248)
        If k(i).Typ = "Verteiler" Then node.Fill.ForeColor.RGB = RGB(255, 238, 190)
        If k(i).Typ = "Sammelkollektor" Then node.Fill.ForeColor.RGB = RGB(213, 239, 210)
        If k(i).DruckFest Then node.Line.Weight = 2
    Next i
    For i = 1 To UBound(e)
        If e(i).Aktiv Then
            a = e(i).Von: b = e(i).Nach
            Set verbindung = ws.Shapes.AddConnector(msoConnectorStraight, 0, 0, 10, 10)
            verbindung.Name = "net_edge_" & i
            ' Keep explicit endpoints; do not let Excel reroute branch lines over other nodes.
            If xpos(b) > xpos(a) Then
                verbindung.ConnectorFormat.BeginConnect ws.Shapes("net_node_" & a), 4
                verbindung.ConnectorFormat.EndConnect ws.Shapes("net_node_" & b), 2
            ElseIf xpos(b) < xpos(a) Then
                verbindung.ConnectorFormat.BeginConnect ws.Shapes("net_node_" & a), 2
                verbindung.ConnectorFormat.EndConnect ws.Shapes("net_node_" & b), 4
            ElseIf ypos(b) >= ypos(a) Then
                verbindung.ConnectorFormat.BeginConnect ws.Shapes("net_node_" & a), 3
                verbindung.ConnectorFormat.EndConnect ws.Shapes("net_node_" & b), 1
            Else
                verbindung.ConnectorFormat.BeginConnect ws.Shapes("net_node_" & a), 1
                verbindung.ConnectorFormat.EndConnect ws.Shapes("net_node_" & b), 3
            End If
            If r.Q(i) > 0.000001 Then verbindung.Line.EndArrowheadStyle = msoArrowheadTriangle
            If r.Q(i) < -0.000001 Then verbindung.Line.BeginArrowheadStyle = msoArrowheadTriangle
            verbindung.Line.ForeColor.RGB = RGB(55, 95, 115)
            verbindung.ZOrder msoSendToBack
            duplicate = 0
            For j = 1 To i - 1
                If (e(j).Von = a And e(j).Nach = b) Or (e(j).Von = b And e(j).Nach = a) Then duplicate = duplicate + 1
            Next j
            sx = 10 + ((xpos(a) + xpos(b) + 140) / 2 - 75) * layoutFaktor
            sy = oben + ((ypos(a) + ypos(b)) / 2 + 0.5 + 24 * duplicate) * layoutFaktor
            Set beschriftung = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, sx, sy, 150 * layoutFaktor, 30 * layoutFaktor)
            beschriftung.Name = "net_label_" & i: beschriftung.TextFrame.Characters.Text = e(i).ID & ": " & k(a).ID & " -> " & k(b).ID & vbLf & "Q=" & Format(r.Q(i), "0.0000") & " m3/h"
            fontsize = 8 * layoutFaktor: If fontsize < 1 Then fontsize = 1
            beschriftung.TextFrame.Characters.Font.Size = fontsize: beschriftung.TextFrame.Characters.Font.Color = vbBlack: beschriftung.Line.Visible = msoFalse
            beschriftung.Fill.ForeColor.RGB = vbWhite
            If e(i).Pumpe Then
                Set pump = ws.Shapes.AddShape(msoShapeOval, sx, sy + 30 * layoutFaktor, 28 * layoutFaktor, 28 * layoutFaktor)
                pump.Name = "net_pump_" & i: pump.TextFrame.Characters.Text = "P"
                pump.TextFrame.Characters.Font.Color = vbBlack
                pump.Fill.ForeColor.RGB = RGB(80, 180, 145)
            End If
        End If
    Next i
    If maxBreite = 0 Then
        For Each node In ws.Shapes
            If Left(node.Name, 4) = "net_" Then
                If node.Left + node.Width > rechteGrenze Then rechteGrenze = node.Left + node.Width
                If node.Top + node.Height > untereGrenze Then untereGrenze = node.Top + node.Height
            End If
        Next node
        ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(3 + Int(untereGrenze / ws.Rows(1).Height), 2 + Int(rechteGrenze / ws.Columns(1).Width))).Address
    End If
End Sub

Public Sub NetzLayoutBerechnen(ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByRef layer() As Long, ByRef xpos() As Double, ByRef ypos() As Double)
    Dim n As Long, i As Long, j As Long, root As Long, head As Long, tail As Long
    Dim queue() As Long, counts() As Long, slots() As Long, other As Long, maxCount As Long
    n = UBound(k): root = 0
    ReDim layer(1 To n): ReDim xpos(1 To n): ReDim ypos(1 To n)
    ReDim queue(1 To n): ReDim counts(0 To n): ReDim slots(0 To n)
    ' Prefer the declared source, never the last pressure reservoir in the table.
    For i = 1 To n
        layer(i) = -1
        If root = 0 And k(i).Typ = "Einspeisung" Then root = i
    Next i
    If root = 0 Then
        For i = 1 To UBound(e)
            If e(i).Aktiv And e(i).Pumpe Then root = e(i).Von: Exit For
        Next i
    End If
    If root = 0 Then
        For i = 1 To n
            If k(i).DruckFest Then root = i: Exit For
        Next i
    End If
    If root = 0 Then root = 1
    head = 1: tail = 1: queue(1) = root: layer(root) = 0
    Do While head <= tail
        i = queue(head): head = head + 1
        For j = 1 To UBound(e)
            If e(j).Aktiv Then
                other = 0
                If e(j).Von = i Then other = e(j).Nach
                If e(j).Nach = i Then other = e(j).Von
                If other > 0 Then
                    If layer(other) < 0 Then
                        layer(other) = layer(i) + 1: tail = tail + 1: queue(tail) = other
                    End If
                End If
            End If
        Next j
    Loop
    For i = 1 To n
        If layer(i) < 0 Then NetzFehler "Schema: nicht verbundenes Netz."
        counts(layer(i)) = counts(layer(i)) + 1
        If counts(layer(i)) > maxCount Then maxCount = counts(layer(i))
    Next i
    For i = 1 To n
        xpos(i) = layer(i) * 300
        ypos(i) = (slots(layer(i)) + (maxCount - counts(layer(i))) / 2) * 115
        slots(layer(i)) = slots(layer(i)) + 1
        If k(i).PositionManuell Then
            If k(i).X < 0 Or k(i).Y < 0 Then NetzFehler "Manuelle X/Y-Positionen muessen >=0 sein."
            xpos(i) = k(i).X: ypos(i) = k(i).Y
        End If
    Next i
End Sub

Public Sub NetzwerkPositionenSpeichern()
    On Error GoTo Fehler
    Dim ws As Worksheet, t As ListObject, i As Long, node As Shape
    Set ws = ThisWorkbook.Worksheets("Netzwerk_Diagramm")
    Set t = ThisWorkbook.Worksheets("Netzwerk_Knoten").ListObjects("tblNetzKnoten")
    For i = 1 To t.ListRows.Count
        Set node = ws.Shapes("net_node_" & i)
        If node.Left < 10 Or node.Top < 75 Then NetzFehler "Knoten innerhalb des Schemabereichs (X/Y>=0) positionieren."
        If Left(node.TextFrame.Characters.Text, Len(CStr(t.ListRows(i).Range.Cells(1, 1).Value)) + 3) <> CStr(t.ListRows(i).Range.Cells(1, 1).Value) & " | " Then NetzFehler "Knotentabelle seit Schema geaendert. Erst neu berechnen."
    Next i
    For i = 1 To t.ListRows.Count
        Set node = ws.Shapes("net_node_" & i)
        t.ListRows(i).Range.Cells(1, 7).Value = node.Left - 10
        t.ListRows(i).Range.Cells(1, 8).Value = node.Top - 75
    Next i
    MsgBox "Manuelle Positionen gespeichert. Zum automatischen Layout X/Y leeren.", vbInformation
    Exit Sub
Fehler:
    MsgBox Err.Description, vbExclamation
End Sub
