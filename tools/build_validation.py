"""Generate VBA assertions from an independent Python bisection oracle."""
import sys,json
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'tests'))
from reference import CASES,point

def literal(x):return format(x,'.17g').replace('e','E')+'#' if 'e' not in format(x,'.17g') else format(x,'.17g').replace('e','E')

def build():
    fixtures=[{'name':name,'inputs':args,'expected':point(*args)} for name,args in CASES]
    (ROOT/'tests/reference_cases.json').write_text(json.dumps(fixtures,indent=2)+'\n')
    lines=['Attribute VB_Name = "Validierung"','Option Explicit','',
    'Private Sub Nah(ByVal actual As Double, ByVal expected As Double, ByVal label As String)',
    '    If Abs(actual - expected) > 0.00000001 * (1 + Abs(expected)) Then Err.Raise vbObjectError + 100, , label & ": Ist=" & actual & " Soll=" & expected',
    'End Sub','', 'Public Sub ValidierungStarten()', '    On Error GoTo Fehler',
    '    Dim r As Betriebspunkt, count As Long, i As Long, bad As Variant, failures As Long, code As Long']
    fields=['V','Re','Lambda','RohrBar','EinzelBar','HoeheBar','BehaelterBar','GesamtBar']
    for fixture in fixtures:
        lines.append('    r = BerechnePunkt('+', '.join(literal(v) for v in fixture['inputs'])+')')
        for field,value in zip(fields,fixture['expected']):
            lines.append(f'    Nah r.{field}, {literal(value)}, "{fixture["name"]}: {field}"')
        lines.append('    count = count + 1')
    lines += ['    For i = 1 To 8','        On Error Resume Next','        Err.Clear',
      '        Select Case i',
      '        Case 1: r = BerechnePunkt(-1, 0.05, 10, 0, 1000, 0.001, 0, 0, 0)',
      '        Case 2: r = BerechnePunkt(1, 0, 10, 0, 1000, 0.001, 0, 0, 0)',
      '        Case 3: r = BerechnePunkt(1, 0.05, -1, 0, 1000, 0.001, 0, 0, 0)',
      '        Case 4: r = BerechnePunkt(1, 0.05, 10, -1, 1000, 0.001, 0, 0, 0)',
      '        Case 5: r = BerechnePunkt(1, 0.05, 10, 0, 0, 0.001, 0, 0, 0)',
      '        Case 6: r = BerechnePunkt(1, 0.05, 10, 0, 1000, 0, 0, 0, 0)',
      '        Case 7: r = BerechnePunkt(1, 0.05, 10, 0, 1000, 0.001, 0, -1, 0)',
      '        Case 8: r = BerechnePunkt(1, 0.05, 10, 0.003, 1000, 0.001, 0, 0, 0)',
      '        End Select','        code = Err.Number','        On Error GoTo Fehler',
      '        If code = 0 Then Err.Raise vbObjectError + 101, , "Ungueltige Eingabe akzeptiert, Fall " & i',
      '        count = count + 1','    Next i',
      '    MsgBox count & " Rechenkern-Testfaelle bestanden (12 Referenzen, 8 Fehlerfaelle).", vbInformation',
      '    Exit Sub','Fehler:', '    MsgBox "VALIDIERUNG FEHLGESCHLAGEN: " & Err.Description, vbCritical', '    Err.Raise vbObjectError + 102, "Validierung", "Validierung fehlgeschlagen"','End Sub']
    (ROOT/'vba/Validierung.bas').write_text('\n'.join(lines)+'\n')
if __name__=='__main__':build()
