"""Generate real-VBA curve checks from independent analytical references."""
from pathlib import Path
import math
ROOT=Path(__file__).resolve().parents[1]
def build():
    resistance=128*.1*20/(math.pi*.05**4*3600*1e5)
    pumpdp=2*1000/200000*(10/3600/(math.pi*.05**2/4))**2
    src=f'''Attribute VB_Name = "NetzFunktionsTests"
Option Explicit

Private Sub Nah(ByVal actual As Double, ByVal expected As Double)
    If Abs(actual - expected) > 0.000001 Then NetzFehler "Kennlinienreferenz weicht ab."
End Sub

Public Sub NetzwerkFunktionenValidieren()
    On Error GoTo Fehler
    Dim k() As NetzKnoten, e() As NetzLeitung, pq(1 To 1) As Double, ph(1 To 1) As Double
    Dim r As NetzResultat, dp As Double, count As Long, code As Long, i As Long
    Dim layers() As Long, xpos() As Double, ypos() As Double
    ReDim k(1 To 3): ReDim e(1 To 2)
    k(1).ID = "Quelle": k(1).DruckFest = True: k(1).DruckBarA = 3
    k(2).ID = "Mitte": k(3).ID = "Ziel": k(3).DruckFest = True: k(3).DruckBarA = 2
    For i = 1 To 2
        e(i).ID = "L" & i: e(i).Von = i: e(i).Nach = i + 1
        e(i).D = 0.05: e(i).L = 10: e(i).Aktiv = True
    Next i
    NetzKennlinienPunkt k, e, 1, 1, 1000, 0.1, pq, ph, r, dp
    If Not r.Erfolg Then NetzFehler r.Status
    Nah dp, {2+resistance:.17g}#: count = count + 1
    Nah r.Q(1), 1: Nah r.Q(2), 1: count = count + 1
    Nah k(1).DruckBarA, 3
    If Not k(1).DruckFest Or k(1).FlussVorgegeben Then NetzFehler "Kennlinie hat Eingaben veraendert."
    count = count + 1
    NetzKennlinienPunkt k, e, 1, 0, 1000, 0.1, pq, ph, r, dp
    If Not r.Erfolg Then NetzFehler r.Status
    Nah dp, 2: count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    k(1).ID = "Saug": k(2).ID = "Druck"
    k(1).DruckFest = True: k(2).DruckFest = True: k(1).DruckBarA = 2: k(2).DruckBarA = 2
    e(1).ID = "P": e(1).Von = 1: e(1).Nach = 2: e(1).D = 0.05
    e(1).Aktiv = True: e(1).Pumpe = True: e(1).Zeta = 2
    NetzKennlinienPunkt k, e, 2, 10, 1000, 0.001, pq, ph, r, dp
    If Not r.Erfolg Then NetzFehler r.Status
    Nah dp, {pumpdp:.17g}#: Nah r.PumpenQ, 10: count = count + 1
    For i = 1 To 3
        On Error Resume Next
        Err.Clear
        Select Case i
        Case 1: NetzKennlinienPunkt k, e, 1, 10, 1000, 0.001, pq, ph, r, dp
        Case 2: NetzKennlinienPunkt k, e, 2, -1, 1000, 0.001, pq, ph, r, dp
        Case 3:
            e(1).Rueckschlag = True
            NetzKennlinienPunkt k, e, 2, 0, 1000, 0.001, pq, ph, r, dp
        End Select
        code = Err.Number
        On Error GoTo Fehler
        If code = 0 Then NetzFehler "Ungueltige Kennlinienvorgabe akzeptiert."
        count = count + 1
    Next i
    ReDim k(1 To 4): ReDim e(1 To 3)
    k(1).ID = "A": k(1).Typ = "Einspeisung": k(1).DruckFest = True
    k(2).ID = "M": k(2).Typ = "Verteiler"
    k(3).ID = "B": k(3).DruckFest = True
    k(4).ID = "S": k(4).DruckFest = True
    e(1).Von = 1: e(1).Nach = 2: e(1).Aktiv = True
    e(2).Von = 2: e(2).Nach = 3: e(2).Aktiv = True
    e(3).Von = 2: e(3).Nach = 4: e(3).Aktiv = True
    NetzLayoutBerechnen k, e, layers, xpos, ypos
    If layers(1) <> 0 Or layers(2) <> 1 Or layers(3) <> 2 Or layers(4) <> 2 Then NetzFehler "Verzweigungslayout startet nicht bei A."
    Nah ypos(2), (ypos(3) + ypos(4)) / 2
    count = count + 1
    k(1).Typ = "": k(4).Typ = "Einspeisung"
    NetzLayoutBerechnen k, e, layers, xpos, ypos
    If layers(4) <> 0 Then NetzFehler "Deklarierte Einspeisung wurde ignoriert."
    count = count + 1
    k(1).PositionManuell = True: k(1).X = 17: k(1).Y = 23
    NetzLayoutBerechnen k, e, layers, xpos, ypos
    Nah xpos(1), 17: Nah ypos(1), 23
    count = count + 1
    MsgBox count & " Netzwerk-Funktionstests bestanden.", vbInformation
    Exit Sub
Fehler:
    MsgBox "FUNKTIONSVALIDIERUNG FEHLGESCHLAGEN: " & Err.Description, vbCritical
    Err.Raise vbObjectError + 260, "NetzFunktionsTests", "Validierung fehlgeschlagen"
End Sub
'''
    (ROOT/'vba/network/NetzFunktionsTests.bas').write_text(src,encoding='ascii')
if __name__=='__main__':build()
