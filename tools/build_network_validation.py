"""Actual-VBA fixtures from independent scalar inversion/nodal oracle.
Does not execute VBA. Does not modify V1 fixtures/modules.
"""
from pathlib import Path
import json,sys,math
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT));sys.path.insert(0,str(ROOT/'tests'))
from network_cases import cases
from network_oracle import oracle
from tools.network.model import G

def lit(x):return format(x,'.17g').replace('e','E')+'#'

def build():
    lines=['Attribute VB_Name = "NetzValidierung"','Option Explicit','',
    'Private Sub Nah(ByVal a As Double, ByVal b As Double, ByVal txt As String)',
    '    If Abs(a-b) > 0.00002 + 0.000001 * Abs(b) Then NetzFehler "Referenzabweichung: " & txt',
    'End Sub','','Public Sub NetzwerkValidieren()',
    '    Dim k() As NetzKnoten, e() As NetzLeitung, pq(1 To 2) As Double, ph(1 To 2) As Double',
    '    Dim r As NetzResultat, count As Long, i As Long',
    '    On Error GoTo Fehler']
    fixtures=[]
    for name,(nodes,edges) in cases().items():
        mu=.001 if name=='turbulent' else .1
        p,q=oracle(nodes,edges,1000,mu)
        fixtures.append(dict(name=name,pressure=p.tolist(),flow=q.tolist(),rho=1000,mu=mu))
        lines += [f'    ReDim k(1 To {len(nodes)}): ReDim e(1 To {len(edges)})',f'    \' {name}']
        for i,n in enumerate(nodes,1):
            lines += [f'    k({i}).ID = "{n.id}": k({i}).Hoehe = {lit(n.height)}']
            if n.pressure is not None:lines += [f'    k({i}).DruckFest = True: k({i}).DruckBarA = {lit(n.pressure)}']
            if n.external is not None:lines += [f'    k({i}).FlussVorgegeben = True: k({i}).ExternQ = {lit(n.external)}']
        for i,e in enumerate(edges,1):
            lines += [f'    e({i}).ID = "{e.id}": e({i}).Von = {e.start+1}: e({i}).Nach = {e.end+1}',
            f'    e({i}).D = {lit(e.d)}: e({i}).L = {lit(e.length)}: e({i}).Eps = {lit(e.rough)}',
            f'    e({i}).Zeta = {lit(e.zeta)}: e({i}).ZKv = {lit(e.zkv)}: e({i}).ZBlende = {lit(e.zorifice)}',
            f'    e({i}).Aktiv = {str(e.active)}: e({i}).Rueckschlag = {str(e.check)}']
        lines += [f'    NetzLoesen k, e, 1000, {lit(mu)}, pq, ph, r',f'    If Not r.Erfolg Then NetzFehler "{name}: " & r.Status']
        for i,value in enumerate(p,1):lines += [f'    Nah r.Druck({i}), {lit(value)}, "{name} Druck {i}"']
        for i,value in enumerate(q,1):lines += [f'    Nah r.Q({i}), {lit(value)}, "{name} Q {i}"']
        lines += ['    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."','    count = count + 1']
    kp=2*1000/2/1e5/(3600*math.pi*.05**2/4)**2
    qp=(-.1+math.sqrt(.01+16*kp))/(2*kp)
    lines += ['    ReDim k(1 To 2): ReDim e(1 To 1)',
    '    k(1).ID = "A": k(2).ID = "B": k(1).DruckFest = True: k(2).DruckFest = True',
    '    k(1).DruckBarA = 2: k(2).DruckBarA = 2',
    '    e(1).ID = "P": e(1).Von = 1: e(1).Nach = 2: e(1).D = 0.05: e(1).Zeta = 2',
    '    e(1).Aktiv = True: e(1).Pumpe = True',
    '    pq(1) = 0: pq(2) = 40: ph(1) = 4 * 100000# / (1000 * G): ph(2) = 0',
    '    NetzLoesen k, e, 1000, 0.001, pq, ph, r',
    '    If Not r.Erfolg Then NetzFehler "Pumpenkopplung: " & r.Status',
    f'    Nah r.Q(1), {lit(qp)}, "Pumpenbetriebspunkt analytisch"','    count = count + 1',
    '    e(1).Pumpe = False: e(1).L = 10: e(1).Zeta = 0',
    '    k(1).DruckFest = False: k(2).DruckFest = False',
    '    NetzLoesen k, e, 1000, 0.1, pq, ph, r',
    '    If r.Erfolg Or InStr(r.Status, "Druckrandbedingung") = 0 Then NetzFehler "Fehlende Druckvorgabe nicht erkannt."','    count = count + 1',
    '    k(1).DruckFest = True: k(1).FlussVorgegeben = True',
    '    NetzLoesen k, e, 1000, 0.1, pq, ph, r',
    '    If r.Erfolg Or InStr(r.Status, "Ueberbestimmt") = 0 Then NetzFehler "Ueberbestimmung nicht erkannt."','    count = count + 1',
    '    k(1).FlussVorgegeben = False: ReDim Preserve k(1 To 3): k(3).ID = "isoliert"',
    '    NetzLoesen k, e, 1000, 0.1, pq, ph, r',
    '    If r.Erfolg Or InStr(r.Status, "Teilnetz") = 0 Then NetzFehler "Insel nicht erkannt."','    count = count + 1',
    '    ReDim Preserve k(1 To 2): k(2).DruckFest = True: k(1).DruckBarA = 3: k(2).DruckBarA = 2.98',
    '    NetzLoesen k, e, 1000, 0.1, pq, ph, r, 1, 0.00000000000001, 0.00000000000001',
    '    If r.Erfolg Then NetzFehler "Iterationslimit wurde ignoriert."','    count = count + 1',
    '    e(1).L = 0',
    '    NetzLoesen k, e, 1000, 0.1, pq, ph, r',
    '    If r.Erfolg Then NetzFehler "Ideale widerstandslose Leitung akzeptiert."','    count = count + 1',
    '    NetzArmaturAddieren e(1), "Zeta", "Frei definierbarer Einzelwiderstand", 1, 1000, 0',
    '    Nah e(1).Zeta, 0, "Expliziter Nullwert bleibt gueltig"','    count = count + 1',
    '    NetzwerkFunktionenValidieren',
    '    MsgBox count & " Netzwerk-Testfaelle bestanden.", vbInformation','    Exit Sub',
    'Fehler:','    MsgBox "NETZVALIDIERUNG FEHLGESCHLAGEN: " & Err.Description, vbCritical',
    '    Err.Raise vbObjectError + 250, "NetzValidierung", "Validierung fehlgeschlagen"','End Sub']
    (ROOT/'vba/network/NetzValidierung.bas').write_text('\n'.join(lines)+'\n',encoding='ascii')
    (ROOT/'tests/network_fixtures.json').write_text(json.dumps(fixtures,indent=2)+'\n')
if __name__=='__main__':build()
