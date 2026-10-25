"""Build actual-VBA tests using independent pressure-recovery and analytic references."""
from pathlib import Path
import sys,math
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tests'))
from reference import orifice_reference, point

def lit(x):return format(x,'.17g').replace('e','E')+'#'

def build():
    lines=['Attribute VB_Name = "ErweiterungsTests"','Option Explicit','',
    'Private Sub Nah(ByVal actual As Double, ByVal expected As Double, ByVal label As String)',
    '    If Abs(actual - expected) > 0.000001 * (1 + Abs(expected)) Then Err.Raise vbObjectError + 150, , label & ": Ist=" & actual & " Soll=" & expected',
    'End Sub','','Public Sub ErweiterungenValidieren()', '    On Error GoTo Fehler',
    '    Dim a As Anlagendaten, r As Betriebspunkt, qs(1 To 3) As Double, hs(1 To 3) As Double',
    '    Dim count As Long, i As Long, code As Long, wert As Double']
    for b in [.2,.25,.5,.75]:
        for cd in [.55,.61,.7]:
            lines += [f'    Nah BlendenZeta(0.05#, {lit(.05*b)}, {lit(cd)}), {lit(orifice_reference(.05,.05*b,cd))}, "Blende beta={b} Cd={cd}"','    count = count + 1']
    lines += ['    Nah KvDruck(10, 10, 1000), 1, "Kv Wasser"', '    count = count + 1',
    '    Nah KvDruck(20, 10, 800), 3.2, "Kv Dichte"', '    count = count + 1',
    '    Nah KvDruck(0, 10, 1000), 0, "Kv Q=0"', '    count = count + 1',
    '    Nah 3 * KvDruck(10, 10, 1000), 3, "Kv in Serie"', '    count = count + 1',
    '    a.D = 0.05: a.Rho = 1000: a.Mu = 0.001',
    '    a.ZKv = KvDruck(1, 10, 1000) * 200000# / 1000# * (3600# * 4 * Atn(1#) * a.D ^ 2 / 4) ^ 2',
    '    a.Z = a.ZKv: r = Anlagenpunkt(10, a)',
    '    Nah r.GesamtBar, 1, "Kv im Anlagenpunkt"', '    count = count + 1',
    '    a.ZKv = 0']
    z=orifice_reference(.05,.025,.61)
    loss=(z+2)*1000/2/1e5*(10/3600/(math.pi*.05**2/4))**2
    lines += ['    a.D = 0.05: a.Rho = 1000: a.Mu = 0.001',
    '    a.BlendeD = 0.05: a.ZBlende = BlendenZeta(0.05, 0.025, 0.61): a.Z = 2 + a.ZBlende',
    '    r = Anlagenpunkt(10, a)',
    f'    Nah r.GesamtBar, {lit(loss)}, "Anlagenpunkt inkl. Blende"',
    '    count = count + 1',
    '    r = Anlagenpunkt(0, a): Nah r.GesamtBar, 0, "Blende Q=0"',
    '    count = count + 1',
    '    a.ZBlende = 0: a.BlendeD = 0: a.Z = 2',
    '    qs(1) = 0: qs(2) = 20: qs(3) = 40',
    '    hs(1) = 4 / (a.Rho * G / 100000#): hs(2) = hs(1) / 2: hs(3) = 0',
    '    r = PumpenSchnittpunkt(a, qs, hs)']
    k=2*1000/2/1e5/(3600*math.pi*.05**2/4)**2
    expected=(-.1+math.sqrt(.01+16*k))/(2*k)
    qc=2300*.001*math.pi*.05/(4*1000)*3600
    below=point(qc*(1-1e-10),.05,100000,0,1000,.001,0,0,0)[-1]
    above=point(qc*(1+1e-10),.05,100000,0,1000,.001,0,0,0)[-1]
    target=(below+above)/2
    lines += [f'    Nah r.Q, {lit(expected)}, "Analytischer Pumpenschnittpunkt"','    count = count + 1',
    '    Nah Pumpendruck(r.Q, qs, hs, a.Rho), r.GesamtBar, "Druckresiduum"','    count = count + 1',
    '    a.Z = 0: a.DP = 1',
    '    hs(1) = 30: hs(2) = 20: hs(3) = 0: qs(2) = 10: qs(3) = 20',
    '    r = PumpenSchnittpunkt(a, qs, hs)',
    f'    Nah r.Q, {lit(20-1e5/(1000*9.80665)/2)}, "Stueckweise lineare Pumpe"','    count = count + 1',
    '    a.DP = 1000 * G * hs(1) / 100000#: r = PumpenSchnittpunkt(a, qs, hs)',
    '    Nah r.Q, 0, "Linker Randpunkt"','    count = count + 1',
    '    a.DP = 0: r = PumpenSchnittpunkt(a, qs, hs)',
    '    Nah r.Q, 20, "Rechter Randpunkt"','    count = count + 1',
    '    For i = 1 To 13','        On Error Resume Next','        Err.Clear','        Select Case i',
    '        Case 1: wert = BlendenZeta(0.05, 0.05, 0.61)',
    '        Case 2: wert = BlendenZeta(0.05, 0.001, 0.61)',
    '        Case 3: wert = BlendenZeta(0.05, 0.025, 0)',
    '        Case 4: wert = Pumpendruck(-1, qs, hs, 1000)',
    '        Case 5: wert = Pumpendruck(21, qs, hs, 1000)',
    '        Case 6: a.DP = 100: r = PumpenSchnittpunkt(a, qs, hs)',
    '        Case 7: a.DP = -1: r = PumpenSchnittpunkt(a, qs, hs)',
    '        Case 8: qs(2) = 0: wert = Pumpendruck(1, qs, hs, 1000)',
    '        Case 9: qs(2) = 10: hs(2) = 40: wert = Pumpendruck(1, qs, hs, 1000)',
    '        Case 10:',
    '            a.L = 100000: a.DP = 0: a.Z = 0',
    f'            qs(1) = 0.3: qs(2) = {lit(qc)}: qs(3) = 0.35',
    f'            hs(1) = {lit((target+.001)*1e5/(1000*9.80665))}',
    f'            hs(2) = {lit(target*1e5/(1000*9.80665))}',
    f'            hs(3) = {lit((target-.001)*1e5/(1000*9.80665))}',
    '            r = PumpenSchnittpunkt(a, qs, hs)',
    '        Case 11: wert = KvDruck(-1, 10, 1000)',
    '        Case 12: wert = KvDruck(1, 0, 1000)',
    '        Case 13: wert = KvDruck(1, 10, 0)',
    '        End Select','        code = Err.Number','        On Error GoTo Fehler',
    '        If code = 0 Then Err.Raise vbObjectError + 151, , "Fehlerfall akzeptiert: " & i',
    '        count = count + 1','    Next i',
    '    MsgBox count & " Erweiterungs-Testfaelle bestanden.", vbInformation',
    '    Exit Sub','Fehler:',
    '    MsgBox "ERWEITERUNGS-VALIDIERUNG FEHLGESCHLAGEN: " & Err.Description, vbCritical',
    '    Err.Raise vbObjectError + 152, "ErweiterungenValidieren", "Validierung fehlgeschlagen"', 'End Sub']
    (ROOT/'vba/ErweiterungsTests.bas').write_text('\n'.join(lines)+'\n',encoding='ascii')
if __name__=='__main__':build()
