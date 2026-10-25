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

class KvStandalone(unittest.TestCase):
    def setUp(self):
        self.z=ZipFile(ROOT/'excel/Kv_Standalone.xlsx');self.addCleanup(self.z.close)
        self.sheet=ET.fromstring(self.z.read('xl/worksheets/sheet1.xml'))
        self.formulas={c.attrib['r']:c.find('s:f',NS).text for c in self.sheet.findall('.//s:c',NS) if c.find('s:f',NS) is not None}
    def test_independent_file_and_embedded_sheet_match(self):
        self.assertIsNone(self.z.testzip())
        wb=ET.fromstring(self.z.read('xl/workbook.xml'))
        self.assertEqual([s.attrib['name'] for s in wb.findall('s:sheets/s:sheet',NS)],['Kv_Standalone'])
        self.assertTrue(all('!' not in f for f in self.formulas.values()))
        with ZipFile(ROOT/'excel/Druckverlust_Vorlage.xlsx') as full:
            embedded=ET.fromstring(full.read('xl/worksheets/sheet10.xml'))
            def contents(sheet,package):
                strings=ET.fromstring(package.read('xl/sharedStrings.xml'))
                texts=[''.join(el.itertext()) for el in strings.findall('s:si',NS)]
                cells={}
                for c in sheet.findall('.//s:c',NS):
                    f=c.find('s:f',NS);v=c.find('s:v',NS)
                    value=v.text if v is not None else None
                    if c.attrib.get('t')=='s':value=texts[int(value)]
                    cells[c.attrib['r']]=(f.text if f is not None else None,value)
                return cells
            self.assertEqual(contents(embedded,full),contents(self.sheet,self.z))
    def test_actual_arithmetic_formulas_against_reference(self):
        # Evaluate only arithmetic expressions from generated XLSX, not Excel functions.
        import ast
        import operator
        ops={ast.Add:operator.add,ast.Sub:operator.sub,ast.Mult:operator.mul,ast.Div:operator.truediv,ast.Pow:operator.pow}
        def calculate(node,values):
            if isinstance(node,ast.Constant):return node.value
            if isinstance(node,ast.Name):return values[node.id]
            if isinstance(node,ast.BinOp):return ops[type(node.op)](calculate(node.left,values),calculate(node.right,values))
            raise AssertionError('Unexpected arithmetic formula')
        for q,kv,rho,n in [(10,10,1000,1),(20,10,800,3),(0,5,1000,2),(1.5,4,1200,1)]:
            values={'D8':q,'D9':kv,'D10':rho,'D11':n}
            for address in ['D21','D22','D23','D24']:
                formula=self.formulas[address]
                self.assertTrue(formula.startswith('IF($D$16=1,') and formula.endswith(',NA())'))
                expr=formula[len('IF($D$16=1,'):-len(',NA())')].replace('^','**')
                values[address]=calculate(ast.parse(expr,mode='eval').body,values)
            expected=n*rho/1000*(q/kv)**2
            self.assertAlmostEqual(values['D22'],expected)
            self.assertAlmostEqual(values['D23'],expected*1e5)
            self.assertAlmostEqual(values['D24'],expected*1e5/(rho*9.80665))
    def test_validation_and_documented_inputs(self):
        check=self.formulas['D16']
        for condition in ['D8>=0','D9>0','D10>0','D11=INT(D11)','D11<=1000','D12>0','ISNUMBER(D9)','D13>=-273.15']:
            self.assertIn(condition,check)
        self.assertTrue(check.startswith('IFERROR('))
        self.assertIn('D16<>1',self.formulas['A17'])
    def test_one_page_a4_with_equations(self):
        setup=self.sheet.find('s:pageSetup',NS)
        self.assertEqual(setup.attrib['paperSize'],'9')
        self.assertEqual(setup.attrib['orientation'],'portrait')
        self.assertEqual(setup.attrib.get('fitToHeight','1'),'1')
        wb=ET.fromstring(self.z.read('xl/workbook.xml'))
        self.assertTrue(any('$A$1:$H$48' in el.text for el in wb.findall('s:definedNames/s:definedName',NS)))
        strings=ET.fromstring(self.z.read('xl/sharedStrings.xml'))
        text=' '.join(el.text or '' for el in strings.findall('.//s:t',NS))
        self.assertIn('Δpeinzeln [bar]',text)
        self.assertIn('ANNAHMEN UND GRENZEN',text)
