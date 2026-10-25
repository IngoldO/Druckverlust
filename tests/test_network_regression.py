import hashlib,json,re
from pathlib import Path
import unittest
from zipfile import ZipFile
import xml.etree.ElementTree as E
ROOT=Path(__file__).resolve().parents[1]
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

class NetworkRegression(unittest.TestCase):
    def test_v1_files_byte_preserved(self):
        baseline=json.loads((ROOT/'docs/v2/Bestand_V1.json').read_text())
        for name,digest in baseline['files'].items():
            with self.subTest(file=name):
                self.assertEqual(hashlib.sha256((ROOT/name).read_bytes()).hexdigest(),digest)
    def test_v2_legacy_sheet_semantics_preserved(self):
        def texts(z):
            return [''.join(s.itertext()) for s in E.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',NS)]
        def cells(z,path):
            strings=texts(z);out={}
            root=E.fromstring(z.read(path))
            for c in root.findall('.//s:c',NS):
                f=c.find('s:f',NS);v=c.find('s:v',NS)
                v=v.text if v is not None else None
                if c.attrib.get('t')=='s':v=strings[int(v)]
                out[c.attrib['r']]=(f.text if f is not None else None,v,c.attrib.get('s'))
            return out
        with ZipFile(ROOT/'excel/Druckverlust_Vorlage.xlsx') as old,ZipFile(ROOT/'excel/Druckverlust_Netzwerk_V2.xlsx') as new:
            self.assertIsNone(new.testzip())
            for i in range(1,11):
                with self.subTest(sheet=i):self.assertEqual(cells(old,f'xl/worksheets/sheet{i}.xml'),cells(new,f'xl/worksheets/sheet{i}.xml'))
            for name in old.namelist():
                if name.startswith(('xl/charts/','xl/tables/')):
                    self.assertEqual(old.read(name),new.read(name))
    def test_network_tables_and_names(self):
        with ZipFile(ROOT/'excel/Druckverlust_Netzwerk_V2.xlsx') as z:
            wb=E.fromstring(z.read('xl/workbook.xml'))
            names=[s.attrib['name'] for s in wb.findall('s:sheets/s:sheet',NS)]
            self.assertEqual(names[-6:],['Netzwerk_Setup','Netzwerk_Knoten','Netzwerk_Leitungen','Netzwerk_Armaturen','Netzwerk_Pumpe','Netzwerk_Ergebnisse'])
            tables=[E.fromstring(z.read(n)).attrib['name'] for n in z.namelist() if n.startswith('xl/tables/')]
            self.assertTrue({'tblNetzKnoten','tblNetzLeitungen','tblNetzArmaturen','tblNetzPumpe'}.issubset(tables))
            defined={n.attrib['name']:n.text for n in wb.findall('s:definedNames/s:definedName',NS)}
            self.assertEqual(defined['NetzKnotenIDs'],'tblNetzKnoten[ID]')
    def test_new_vba_importability_and_entrypoints(self):
        paths=list((ROOT/'vba/network').glob('*.bas'))
        self.assertEqual(len(paths),5)
        procedures=[]
        for path in paths:
            data=path.read_text(encoding='ascii')
            self.assertIn('Option Explicit',data)
            self.assertLess(max(map(len,data.splitlines())),1024)
            procedures+=re.findall(r'(?im)^Public (?:Sub|Function)\s+(\w+)',data)
        self.assertEqual(len(procedures),len(set(procedures)))
        self.assertTrue({'NetzwerkBerechnen','NetzwerkValidieren','NetzLoesen','NetzArmaturAddieren'}.issubset(procedures))
    def test_fixtures_and_generated_validation(self):
        fixtures=json.loads((ROOT/'tests/network_fixtures.json').read_text())
        self.assertEqual(len(fixtures),18)
        large=next(f for f in fixtures if f['name']=='mesh20x30')
        self.assertEqual(len(large['pressure']),20);self.assertEqual(len(large['flow']),30)
        src=(ROOT/'vba/network/NetzValidierung.bas').read_text()
        self.assertEqual(src.count('    count = count + 1'),25)
