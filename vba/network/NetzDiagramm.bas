Attribute VB_Name = "NetzDiagramm"
Option Explicit

Public Sub NetzSchemaZeichnen(ByVal ws As Worksheet, ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByRef r As NetzResultat, Optional ByVal oben As Double = 75, Optional ByVal maxBreite As Double = 0, Optional ByVal maxHoehe As Double = 0)
    Dim layer() As Long, counts() As Long, xpos() As Double, ypos() As Double
    Dim i As Long, j As Long, a As Long, b As Long, root As Long, pass As Long, n As Long
    Dim changed As Boolean, width As Double, height As Double, scale As Double, sx As Double
    Dim right As Double, bottom As Double, fontsize As Double
    Dim sy As Double, duplicate As Long, node As Shape, line As Shape, label As Shape, pump As Shape
    If Not r.Erfolg Then NetzFehler "Kein gueltiger Betriebspunkt fuer das Schema."
    NetzShapesLoeschen ws
    n = UBound(k): ReDim layer(1 To n): ReDim counts(0 To n): ReDim xpos(1 To n): ReDim ypos(1 To n)
    root = 1
    For i = 1 To n
        layer(i) = -1
        If k(i).DruckFest Then root = i
    Next i
    layer(root) = 0
    For pass = 1 To n
        changed = False
        For i = 1 To UBound(e)
            If e(i).Aktiv Then
                a = e(i).Von: b = e(i).Nach
                If layer(a) >= 0 And layer(b) < 0 Then layer(b) = layer(a) + 1: changed = True
                If layer(b) >= 0 And layer(a) < 0 Then layer(a) = layer(b) + 1: changed = True
            End If
        Next i
        If Not changed Then Exit For
    Next pass
    For i = 1 To n
        If layer(i) < 0 Then NetzFehler "Schema: nicht verbundenes Netz."
        xpos(i) = layer(i) * 200: ypos(i) = counts(layer(i)) * 115
        counts(layer(i)) = counts(layer(i)) + 1
        If k(i).PositionManuell Then
            If k(i).X < 0 Or k(i).Y < 0 Then NetzFehler "Manuelle X/Y-Positionen muessen >=0 sein."
            xpos(i) = k(i).X: ypos(i) = k(i).Y
        End If
        If xpos(i) + 140 > width Then width = xpos(i) + 140
        If ypos(i) + 65 > height Then height = ypos(i) + 65
    Next i
    ' Include labels/pump symbols in the fitting bounds, also for parallel edges.
    For i = 1 To UBound(e)
        If e(i).Aktiv Then
            a = e(i).Von: b = e(i).Nach: duplicate = 0
            For j = 1 To i - 1
                If (e(j).Von = a And e(j).Nach = b) Or (e(j).Von = b And e(j).Nach = a) Then duplicate = duplicate + 1
            Next j
            sx = (xpos(a) + xpos(b) + 140) / 2 + 150
            sy = (ypos(a) + ypos(b)) / 2 + 17 * duplicate + 60
            If sx > width Then width = sx
            If sy > height Then height = sy
        End If
    Next i
    scale = 1
    If maxBreite > 0 And width > maxBreite Then scale = maxBreite / width
    If maxHoehe > 0 And height * scale > maxHoehe Then scale = maxHoehe / height
    For i = 1 To n
        sx = 10 + xpos(i) * scale: sy = oben + ypos(i) * scale
        Set node = ws.Shapes.AddShape(msoShapeRoundedRectangle, sx, sy, 140 * scale, 65 * scale)
        node.Name = "net_node_" & i
        node.TextFrame.Characters.Text = k(i).ID & " | " & k(i).Bezeichnung & vbLf & k(i).Typ & vbLf & Format(r.Druck(i), "0.0000") & " bar(a)"
        fontsize = 9 * scale: If fontsize < 1 Then fontsize = 1
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
            Set line = ws.Shapes.AddConnector(msoConnectorElbow, 0, 0, 10, 10)
            line.Name = "net_edge_" & i
            line.ConnectorFormat.BeginConnect ws.Shapes("net_node_" & a), 4
            line.ConnectorFormat.EndConnect ws.Shapes("net_node_" & b), 2
            line.RerouteConnections
            If r.Q(i) > 0.000001 Then line.Line.EndArrowheadStyle = msoArrowheadTriangle
            If r.Q(i) < -0.000001 Then line.Line.BeginArrowheadStyle = msoArrowheadTriangle
            line.Line.ForeColor.RGB = RGB(55, 95, 115)
            line.ZOrder msoSendToBack
            duplicate = 0
            For j = 1 To i - 1
                If (e(j).Von = a And e(j).Nach = b) Or (e(j).Von = b And e(j).Nach = a) Then duplicate = duplicate + 1
            Next j
            sx = 10 + (xpos(a) + xpos(b) + 140) * scale / 2
            sy = oben + (ypos(a) + ypos(b)) * scale / 2 + 17 * duplicate * scale
            Set label = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, sx, sy, 150 * scale, 30 * scale)
            label.Name = "net_label_" & i: label.TextFrame.Characters.Text = e(i).ID & " | Q=" & Format(r.Q(i), "0.0000") & " m3/h"
            fontsize = 8 * scale: If fontsize < 1 Then fontsize = 1
            label.TextFrame.Characters.Font.Size = fontsize: label.TextFrame.Characters.Font.Color = vbBlack: label.Line.Visible = msoFalse
            label.Fill.ForeColor.RGB = vbWhite
            If e(i).Pumpe Then
                Set pump = ws.Shapes.AddShape(msoShapeOval, sx, sy + 30 * scale, 28 * scale, 28 * scale)
                pump.Name = "net_pump_" & i: pump.TextFrame.Characters.Text = "P"
                pump.TextFrame.Characters.Font.Color = vbBlack
                pump.Fill.ForeColor.RGB = RGB(80, 180, 145)
            End If
        End If
    Next i
    If maxBreite = 0 Then
        For Each node In ws.Shapes
            If Left(node.Name, 4) = "net_" Then
                If node.Left + node.Width > right Then right = node.Left + node.Width
                If node.Top + node.Height > bottom Then bottom = node.Top + node.Height
            End If
        Next node
        ws.PageSetup.PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(3 + Int(bottom / ws.Rows(1).Height), 2 + Int(right / ws.Columns(1).Width))).Address
    End If
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
