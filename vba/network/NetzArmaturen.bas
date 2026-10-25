Attribute VB_Name = "NetzArmaturen"
Option Explicit

' Assign exactly ONE model per fitting; a numeric zero is an explicit override.
Public Sub NetzArmaturAddieren(ByRef e As NetzLeitung, ByVal modell As String, ByVal typ As String, ByVal n As Double, ByVal rho As Double, Optional ByVal zetaManuell As Variant, Optional ByVal kv As Double = 0, Optional ByVal lokalD As Double = 0, Optional ByVal lochD As Double = 0, Optional ByVal cd As Double = 0)
    Dim z As Double
    If n < 0 Or n <> Fix(n) Or n > 100000 Then NetzFehler "Armaturenanzahl muss ganzzahlig zwischen 0 und 100000 liegen."
    If n = 0 Then Exit Sub
    Select Case modell
    Case "Zeta", "Rueckschlag"
        If IsMissing(zetaManuell) Or IsEmpty(zetaManuell) Then
            If typ = "Frei definierbarer Einzelwiderstand" Then NetzFehler "Freier Widerstand benoetigt expliziten Zeta-Wert."
            z = StandardZeta(typ)
        Else
            If IsError(zetaManuell) Then NetzFehler "Zellfehler im Zeta-Wert."
            If VarType(zetaManuell) = vbBoolean Or Not IsNumeric(zetaManuell) Then NetzFehler "Zeta muss numerisch sein."
            z = CDbl(zetaManuell)
        End If
        If z < 0 Then NetzFehler "Zeta muss >=0 sein."
        e.Zeta = e.Zeta + n * z
        If modell = "Rueckschlag" Or typ = "R" & ChrW(252) & "ckschlagventil" Then e.Rueckschlag = True
    Case "Kv"
        e.ZKv = e.ZKv + NetzKvZeta(kv, n, e.D, rho)
    Case "Blende"
        e.ZBlende = e.ZBlende + NetzBlendenZeta(lokalD, lochD, cd, n, e.D)
    Case Else
        NetzFehler "Unbekanntes Armaturenmodell: " & modell
    End Select
End Sub
