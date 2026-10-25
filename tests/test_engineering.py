"""Numerical references and print/formula contracts; Excel runtime remains untested."""
import math
from pathlib import Path
import unittest
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from reference import orifice_reference
ROOT=Path(__file__).resolve().parents[1]
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

class Engineering(unittest.TestCase):
    def setUp(self):
        self.z=ZipFile(ROOT/'excel/Druckverlust_Vorlage.xlsx');self.addCleanup(self.z.close)
    def xml(self,name):return ET.fromstring(self.z.read(name))
    def test_kv_reference_water_and_density(self):
        for q,kv,rho,n,expected in [(10,10,1000,1,1),(20,10,800,1,3.2),(0,10,1000,1,0),(10,10,1000,3,3)]:
            dp=n*rho/1000*(q/kv)**2
            self.assertAlmostEqual(dp,expected)
    def test_equivalent_zeta_preserves_kv_pressure(self):
        for D in [.02,.05,.1]:
            for q in [0,5,20]:
                rho,kv,n=850,12,2
                # Formula implemented in LeseAnlage, compared to direct Kv definition.
                z=n*(rho/1000/kv**2)*200000/rho*(3600*math.pi*D**2/4)**2
                v=q/3600/(math.pi*D**2/4)
                self.assertAlmostEqual(z*rho*v*v/200000,n*rho/1000*(q/kv)**2,12)
    def test_standalone_has_no_other_sheet_dependencies(self):
        sheet=self.xml('xl/worksheets/sheet8.xml')
        formulas=[f.text for f in sheet.findall('.//s:f',NS)]
        self.assertEqual(len(formulas),10)
        self.assertTrue(all('!' not in f for f in formulas))
        self.assertTrue(any('SQRT' in f for f in formulas))
        self.assertTrue(any('NA()' in f for f in formulas))
        self.assertTrue(any('ISNUMBER' in f for f in formulas))
        with ZipFile(ROOT/'excel/Drosselblende_Standalone.xlsx') as single:
            wb=ET.fromstring(single.read('xl/workbook.xml'))
            self.assertEqual(len(wb.findall('s:sheets/s:sheet',NS)),1)
            local=ET.fromstring(single.read('xl/worksheets/sheet1.xml'))
            self.assertEqual([f.text for f in local.findall('.//s:f',NS)],formulas)
            self.assertIsNone(single.testzip())
    def test_standalone_numeric_example_and_recovery(self):
        q,D,d,rho,mu,cd=10,.0545,.025,998.2,.001002,.61
        v=q/3600/(math.pi*D**2/4);beta=d/D
        tap=(1-beta**4)/(cd**2*beta**4)*rho*v*v/200000
        permanent=orifice_reference(D,d,cd)*rho*v*v/200000
        self.assertGreater(tap,permanent)
        self.assertGreater(permanent,0)
        self.assertGreater(rho*v*D/mu,10000)
    def test_a4_print_contracts(self):
        for sheetnum in [8,9]:
            root=self.xml(f'xl/worksheets/sheet{sheetnum}.xml')
            setup=root.find('s:pageSetup',NS)
            self.assertEqual(setup.attrib['paperSize'],'9')
            self.assertEqual(setup.attrib['orientation'],'portrait')
            self.assertEqual(setup.attrib.get('fitToWidth','1'),'1')
        wb=self.xml('xl/workbook.xml')
        areas=[el.text for el in wb.findall('s:definedNames/s:definedName',NS) if el.attrib['name']=='_xlnm.Print_Area']
        self.assertTrue(any('Blende_Standalone' in a and '$A$1:$H$48' in a for a in areas))
        self.assertTrue(any('Systembericht' in a and '$A$1:$H$44' in a for a in areas))
    def test_report_recomputes_inputs_and_copies_chart_data(self):
        src=(ROOT/'vba/EngineeringBericht.bas').read_text()
        self.assertIn('a = LeseAnlage()',src)
        self.assertIn('op = PumpenSchnittpunkt(a, qs, hs)',src)
        self.assertIn('.PlotVisibleOnly = False',src)
        self.assertIn('NICHT AKTUELL',src)
        self.assertIn('ws.Range("N5").Resize(n, 4).Value = daten',src)
        self.assertIn('ThisWorkbook.Worksheets("KvArmatur")',(ROOT/'vba/Eingaben.bas').read_text())
    def test_kv_sheet_and_result_columns(self):
        kv=self.xml('xl/worksheets/sheet7.xml')
        cells={c.attrib['r']:c for c in kv.findall('.//s:c',NS)}
        self.assertEqual(float(cells['B5'].find('s:v',NS).text),0)
        self.assertEqual(float(cells['B6'].find('s:v',NS).text),10)
        for name in ['Oberflaeche','Betriebsoberflaeche']:
            self.assertIn('a.ZKv * a.Rho', (ROOT/f'vba/{name}.bas').read_text())
