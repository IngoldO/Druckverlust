Attribute VB_Name = "Eingaben"
Option Explicit

Public Function Zahl(ByVal c As Range, ByVal bezeichnung As String) As Double
    If IsError(c.Value) Then Err.Raise vbObjectError + 10, , bezeichnung & ": Zellfehler."
    If IsEmpty(c.Value) Or VarType(c.Value) = vbBoolean Then Err.Raise vbObjectError + 11, , bezeichnung & ": Zahl fehlt."
    If Not IsNumeric(c.Value) Then Err.Raise vbObjectError + 12, , bezeichnung & ": Zahl erforderlich."
    Zahl = CDbl(c.Value)
End Function

Public Function StandardZeta(ByVal typ As String) As Double
    Dim db As ListObject, row As ListRow
    Set db = ThisWorkbook.Worksheets("Armaturendatenbank").ListObjects("tblDatenbank")
    For Each row In db.ListRows
        If CStr(row.Range.Cells(1, 1).Value) = typ Then
            StandardZeta = Zahl(row.Range.Cells(1, 2), "Datenbank-Zeta")
            If StandardZeta < 0 Then Err.Raise vbObjectError + 13, , "Datenbank-Zeta muss >=0 sein."
            Exit Function
        End If
    Next row
    Err.Raise vbObjectError + 14, , "Unbekannte Armatur: " & typ
End Function

Public Function SummeZeta() As Double
    Dim t As ListObject, row As ListRow, typ As String, n As Double, z As Double, total As Double
    Set t = ThisWorkbook.Worksheets("Eingabe").ListObjects("tblArmaturen")
    For Each row In t.ListRows
        If IsError(row.Range.Cells(1, 1).Value) Then Err.Raise vbObjectError + 15, , "Zellfehler im Armaturentyp."
        typ = Trim(CStr(row.Range.Cells(1, 1).Value))
        If typ = "" And Application.CountA(row.Range.Cells(1, 2).Resize(1, 2)) = 0 Then GoTo NextRow
        If typ = "" Then Err.Raise vbObjectError + 16, , "Armaturentyp fehlt."
        n = Zahl(row.Range.Cells(1, 2), "Armaturenanzahl")
        If n < 0 Or n <> Fix(n) Or n > 100000 Then Err.Raise vbObjectError + 17, , "Anzahl muss ganzzahlig zwischen 0 und 100000 sein."
        If typ = "Frei definierbarer Einzelwiderstand" And n > 0 And IsEmpty(row.Range.Cells(1, 3).Value) Then Err.Raise vbObjectError + 24, , "Fuer einen aktiven freien Einzelwiderstand Zeta manuell angeben."
        z = StandardZeta(typ)
        If Not IsEmpty(row.Range.Cells(1, 3).Value) Then z = Zahl(row.Range.Cells(1, 3), "Zeta-Ueberschreibung")
        If z < 0 Then Err.Raise vbObjectError + 18, , "Zeta muss >=0 sein."
        total = total + n * z
NextRow:
    Next row
    SummeZeta = total
End Function

Public Sub DNAnwenden()
    On Error GoTo Fehler
    Dim ws As Worksheet, db As Worksheet, i As Long, dn As Double
    Set ws = ThisWorkbook.Worksheets("Eingabe")
    Set db = ThisWorkbook.Worksheets("Armaturendatenbank")
    dn = Zahl(ws.Range("B5"), "DN")
    For i = 5 To 22
        If db.Cells(i, 6).Value = dn Then
            ws.Range("B6").Value = db.Cells(i, 7).Value
            MsgBox "Vorschlag eingesetzt. Tatsaechlichen Innendurchmesser anhand Rohrnorm/Wandstaerke pruefen.", vbInformation
            Exit Sub
        End If
    Next i
    Err.Raise vbObjectError + 19, , "DN nicht in Vorschlagstabelle."
    Exit Sub
Fehler:
    MsgBox Err.Description, vbExclamation, "DN-Vorschlag"
End Sub

Public Sub ArmaturHinzufuegen()
    On Error GoTo Fehler
    Dim r As ListRow
    Set r = ThisWorkbook.Worksheets("Eingabe").ListObjects("tblArmaturen").ListRows.Add
    r.Range.Cells(1, 1).Value = "Frei definierbarer Einzelwiderstand"
    r.Range.Cells(1, 2).Value = 0
    r.Range.Cells(1, 3).Value = 0
    With r.Range.Cells(1, 1).Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, xlBetween, "=ArmaturenTypen"
    End With
    Exit Sub
Fehler:
    MsgBox Err.Description, vbExclamation
End Sub
