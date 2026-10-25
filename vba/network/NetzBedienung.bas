Attribute VB_Name = "NetzBedienung"
Option Explicit

Public Sub NetzToleranzen(ByRef lim As Long, ByRef tq As Double, ByRef tp As Double)
    Dim v As Double
    With ThisWorkbook.Worksheets("Netzwerk_Setup")
        v = Zahl(.Range("B7"), "Iterationslimit")
        tq = Zahl(.Range("B8"), "Volumentoleranz"): tp = Zahl(.Range("B9"), "Drucktoleranz")
    End With
    If v < 1 Or v > 1000 Or v <> Fix(v) Or tq <= 0 Or tp <= 0 Then NetzFehler "Iterationslimit 1..1000, positive Toleranzen erforderlich."
    lim = CLng(v)
End Sub

Public Sub NetzShapesLoeschen(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left(ws.Shapes(i).Name, 4) = "net_" Then ws.Shapes(i).Delete
    Next i
End Sub

Public Sub NetzAusgabeLeeren(ByVal ws As Worksheet)
    NetzShapesLoeschen ws
    ws.Cells.ClearContents
End Sub

Public Sub NetzwerkEinrichtung()
    On Error GoTo Fehler
    Dim ws As Worksheet, b As Shape, i As Long, labels As Variant, macros As Variant
    Set ws = ThisWorkbook.Worksheets("Netzwerk_Setup")
    labels = Array("Netzwerk berechnen + Schema", "Systemkennlinie berechnen", "Engineering-Bericht erstellen", "Knotenpositionen speichern", "Netzwerk validieren")
    macros = Array("NetzwerkBerechnen", "NetzwerkKennlinieBerechnen", "NetzwerkBerichtErstellen", "NetzwerkPositionenSpeichern", "NetzwerkValidieren")
    For i = 0 To UBound(labels)
        On Error Resume Next
        ws.Shapes("net_button_" & i).Delete
        On Error GoTo Fehler
        Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, ws.Range("H5").Left, ws.Range("H5").Top + 45 * i, 260, 34)
        b.Name = "net_button_" & i: b.TextFrame.Characters.Text = labels(i)
        b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macros(i)
        b.Fill.ForeColor.RGB = RGB(24, 84, 105): b.TextFrame.Characters.Font.Color = vbWhite
    Next i
    MsgBox "Netzwerk-Schaltflaechen eingerichtet; Modus A bleibt unveraendert.", vbInformation
    Exit Sub
Fehler:
    MsgBox Err.Description, vbExclamation
End Sub
