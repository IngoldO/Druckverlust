from pathlib import Path
import copy,math,unittest
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import numpy as np
from tools.network.model import Node as N, Edge as E,G
from tools.network.curve import point
from tools.network.newton import solve
from network_oracle import oracle
ROOT=Path(__file__).resolve().parents[1]
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

class Characteristic(unittest.TestCase):
    def test_port_analytic_series_and_no_input_mutation(self):
        nodes=[N('source',3),N('mid'),N('sink',2)]
        edges=[E('a',0,1),E('b',1,2)]
        saved=copy.deepcopy((nodes,edges))
        for q in [0,.1,1,2]:
            p,r=point(nodes,edges,0,q,mu=.1)
            resistance=128*.1*20/(math.pi*.05**4*3600*1e5)
            self.assertAlmostEqual(p,2+resistance*q,7)
            np.testing.assert_allclose(r['q'],[q,q],atol=1e-6)
        self.assertEqual((nodes,edges),saved)
    def test_parallel_recomputes_distribution_and_reference(self):
        nodes=[N('source',3),N('sink',2)]
        edges=[E('a',0,1,d=.03,length=15),E('b',0,1,d=.05,length=60,zeta=8)]
        ratios=[]
        for q in [2,10,30]:
            dp,r=point(nodes,edges,0,q)
            reference_nodes=[N('source',dp),N('sink',2)]
            pp,qq=oracle(reference_nodes,edges,1000,.001)
            np.testing.assert_allclose(r['q'],qq,rtol=2e-6,atol=2e-5)
            self.assertAlmostEqual(sum(r['q']),q,5)
            ratios.append(r['q'][0]/q)
        self.assertGreater(max(ratios)-min(ratios),1e-4)
    def test_pump_curve_analytic_including_zero_and_outside_manufacturer_range(self):
        nodes=[N('suction',2),N('discharge',2)]
        edges=[E('pump',0,1,length=0,zeta=2,pump=True)]
        for q in [0,10,50]:
            dp,r=point(nodes,edges,1,q)
            v=q/3600/(math.pi*.05**2/4)
            self.assertAlmostEqual(dp,2*1000*v*v/200000,7)
            self.assertAlmostEqual(r['q'][0],q,5)
    def test_closed_loop_pump_required_head(self):
        nodes=[N('tank',2),N('discharge'),N('return')]
        edges=[E('pump',0,1,length=0,pump=True),E('line1',1,2),E('line2',2,0)]
        dp,r=point(nodes,edges,1,1,mu=.1)
        resistance=128*.1*20/(math.pi*.05**4*3600*1e5)
        self.assertAlmostEqual(dp,resistance,7)
        np.testing.assert_allclose(r['q'],np.ones(3),atol=1e-6)
    def test_system_curve_at_coupled_operating_point_matches_pump(self):
        nodes=[N('a',2),N('b',2)]
        edges=[E('pump',0,1,length=0,zeta=2,pump=True)]
        curve=([0,40],[4e5/(1000*G),0])
        op=solve(nodes,edges,curve=curve)
        dp,r=point(nodes,edges,1,op['q'][0])
        self.assertAlmostEqual(dp,op['pump_pressure'],7)
    def test_curve_rejects_undefined_boundary_conditions(self):
        with self.assertRaises(ValueError):point([N('a',2),N('b')],[E('l',0,1)],0,1)
        with self.assertRaises(ValueError):point([N('a',2),N('b',3)],[E('p',0,1,pump=True)],0,1)
        with self.assertRaises(ValueError):point([N('a',2),N('b',3),N('c',external=1)],[E('l',0,1),E('m',1,2)],0,1)
    def test_curve_rejects_zero_with_pump_check(self):
        with self.assertRaises(ValueError):point([N('a',2),N('b',2)],[E('p',0,1,pump=True,check=True)],1,0)

class EngineeringFeatures(unittest.TestCase):
    def test_setup_and_output_sheet_contracts(self):
        with ZipFile(ROOT/'excel/Druckverlust_Netzwerk_V2.xlsx') as z:
            wb=ET.fromstring(z.read('xl/workbook.xml'))
            names=[s.attrib['name'] for s in wb.findall('s:sheets/s:sheet',NS)]
            self.assertTrue({'Netzwerk_Diagramm','Netzwerk_Kennlinie','Netzwerk_Bericht'}.issubset(names))
            setup=ET.fromstring(z.read(f'xl/worksheets/sheet{names.index("Netzwerk_Setup")+1}.xml'))
            refs={c.attrib['r'] for c in setup.findall('.//s:c',NS)}
            self.assertTrue({'B17','B18','B19','B20'}.issubset(refs))
            report=ET.fromstring(z.read(f'xl/worksheets/sheet{names.index("Netzwerk_Bericht")+1}.xml'))
            page=report.find('s:pageSetup',NS)
            self.assertEqual(page.attrib['paperSize'],'9');self.assertEqual(page.attrib['orientation'],'landscape')
    def test_drawing_owns_only_prefixed_shapes_and_supports_reverse(self):
        src=(ROOT/'vba/network/NetzDiagramm.bas').read_text()
        self.assertIn('msoConnectorStraight',src)
        self.assertIn('BeginArrowheadStyle',src);self.assertIn('EndArrowheadStyle',src)
        self.assertIn('PositionManuell',src);self.assertIn('NetzwerkPositionenSpeichern',src)
        cleanup=(ROOT/'vba/network/NetzBedienung.bas').read_text()
        self.assertIn('Left(ws.Shapes(i).Name, 4) = "net_"',cleanup)
    def test_report_recomputes_and_preserves_metadata(self):
        src=(ROOT/'vba/network/NetzBericht.bas').read_text()
        self.assertIn('NetzLoesen k, e',src);self.assertIn('NetzSchemaZeichnen ws',src)
        self.assertIn('NetzKennlinienPunkt k, e',src);self.assertIn('.PlotVisibleOnly = False',src)
        self.assertIn('ws.Range("A6:Z" & last).ClearContents',src)
        self.assertIn('NICHT VERFUEGBAR',src);self.assertIn('UNGUELTIG',src)
