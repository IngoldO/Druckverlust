"""Cloud checks of the numerical algorithm and XLSX structure.
These do NOT execute/compile VBA; run ValidierungStarten in Windows Excel too.
"""
import json
import math
from pathlib import Path
import unittest
import zipfile
import xml.etree.ElementTree as ET
from reference import CASES, factor, point
ROOT=Path(__file__).resolve().parents[1]
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def fixed_point(re, rough):
    """Python transcription of VBA iteration for numerical cross-check only."""
    if re == 0:return 0.
    if re < 2300:return 64/re
    x=7.
    for _ in range(200):
        xn=-2*math.log10(rough/3.7+2.51*x/re)
        if abs(xn-x)<1e-10:
            f=1/xn**2
            return max(f,64/re) if re<4000 else f
        x=xn
    raise AssertionError('No convergence')

class NumericalValidation(unittest.TestCase):
    def test_laminar_hagen_poiseuille(self):
        args=CASES[0][1];q,d,l,eps,rho,mu,h,z,dp=args
        result=point(*args)
        expected=128*mu*l*(q/3600)/(math.pi*d**4)/1e5
        self.assertAlmostEqual(result[3],expected,14)
    def test_known_smooth_colebrook(self):
        self.assertAlmostEqual(factor(1e5,0),.01798977308427384,13)
    def test_iteration_against_independent_bisection(self):
        for re in [2300,2301,3000,3999,4000,1e4,1e5,1e6,1e8]:
            for rough in [0,1e-6,.0002,.01,.05]:
                with self.subTest(re=re,rough=rough):
                    self.assertAlmostEqual(fixed_point(re,rough),factor(re,rough),11)
    def test_zero_flow_static_only(self):
        result=point(0,.05,100,.000045,1000,.001,12,3,1.5)
        self.assertEqual(result[:5],[0,0,0,0,0])
        self.assertAlmostEqual(result[-1],2.676798)
    def test_fittings_are_additive(self):
        a=point(10,.05,100,0,1000,.001,0,0,0)
        b=point(10,.05,100,0,1000,.001,0,2*.9+3*.4+10,0)
        self.assertAlmostEqual(b[-1]-a[-1],13*a[0]**2*1000/2/1e5)
    def test_static_sign_and_vessel_pressure(self):
        self.assertAlmostEqual(point(0,.05,0,0,1000,.001,-10,0,-.5)[-1],-1.480665)
    def test_transition_is_explicit_upper_envelope(self):
        self.assertGreaterEqual(factor(3000,0),64/3000)
        self.assertIn('UNSICHER', (ROOT/'vba/Hydraulik.bas').read_text())
        self.assertIn('ElseIf r.re < 4000', (ROOT/'vba/Hydraulik.bas').read_text())
    def test_reference_fixtures_not_stale(self):
        fixtures=json.loads((ROOT/'tests/reference_cases.json').read_text())
        self.assertEqual(len(fixtures),len(CASES))
        for fixture,(name,args) in zip(fixtures,CASES):
            self.assertEqual(fixture['name'],name)
            self.assertEqual(fixture['inputs'],list(args))
            self.assertEqual(fixture['expected'],point(*args))
    def test_vba_selftests_have_references_and_errors(self):
        src=(ROOT/'vba/Validierung.bas').read_text()
        self.assertEqual(src.count('    Nah r.'),12*8)
        self.assertIn('For i = 1 To 8',src)
        self.assertIn('If code = 0 Then Err.Raise',src)

class WorkbookValidation(unittest.TestCase):
    def setUp(self):
        self.z=zipfile.ZipFile(ROOT/'excel/Druckverlust_Vorlage.xlsx')
        self.addCleanup(self.z.close)
    def xml(self,name):return ET.fromstring(self.z.read(name))
    def test_archive_integrity(self):self.assertIsNone(self.z.testzip())
    def test_sheets_and_named_ranges(self):
        root=self.xml('xl/workbook.xml')
        self.assertEqual([s.attrib['name'] for s in root.findall('s:sheets/s:sheet',NS)],['Eingabe','Ergebnisse','Armaturendatenbank','Lochblende','Pumpe','Betriebspunkte','KvArmatur','Blende_Standalone','Systembericht'])
        self.assertEqual({n.attrib['name'] for n in root.findall('s:definedNames/s:definedName',NS) if not n.attrib['name'].startswith('_xlnm.')},{'DNWerte','ArmaturenTypen'})
    def test_table_contract(self):
        tables=[self.xml(n) for n in self.z.namelist() if n.startswith('xl/tables/')]
        self.assertEqual({t.attrib['name'] for t in tables},{'tblArmaturen','tblDatenbank','tblPumpe'})
        self.assertEqual({t.attrib['ref'] for t in tables},{'A20:D32','A4:D16','A20:C26'})
    def test_parameter_cells(self):
        cells={c.attrib['r']:c for c in self.xml('xl/worksheets/sheet1.xml').findall('.//s:c',NS)}
        for i in range(5,17):self.assertIn(f'B{i}',cells)
        self.assertEqual(float(cells['B6'].find('s:v',NS).text),54.5)
        self.assertEqual(float(cells['G3'].find('s:v',NS).text),10)
        orifice={c.attrib['r']:c for c in self.xml('xl/worksheets/sheet4.xml').findall('.//s:c',NS)}
        for address,value in [('B5',0),('B6',50),('B7',54.5),('B8',25),('B9',.61),('B10',1)]:
            self.assertEqual(float(orifice[address].find('s:v',NS).text),value)
        pump_table=self.xml('xl/tables/table3.xml')
        self.assertEqual(pump_table.attrib['ref'],'A20:C26')
        self.assertIn('G3',(ROOT/'vba/Betriebsoberflaeche.bas').read_text())
        self.assertIn('ErweiterungenValidieren',(ROOT/'vba/Validierung.bas').read_text())
    def test_chart_three_xy_series(self):
        ns={'c':'http://schemas.openxmlformats.org/drawingml/2006/chart'}
        chart=self.xml('xl/charts/chart1.xml')
        self.assertEqual(len(chart.findall('.//c:scatterChart/c:ser',ns)),3)
    def test_import_files_are_ascii_and_modules(self):
        files=list((ROOT/'vba').glob('*.bas'))
        self.assertEqual(len(files),8)
        for f in files:
            data=f.read_bytes();data.decode('ascii')
            self.assertIn(b'Option Explicit',data)
            self.assertLess(max(map(len,data.splitlines())),1024)

if __name__=='__main__':unittest.main(verbosity=2)
