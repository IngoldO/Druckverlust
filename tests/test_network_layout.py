from pathlib import Path
import unittest
from tools.network.layout import layout
ROOT=Path(__file__).resolve().parents[1]

class NetworkLayout(unittest.TestCase):
    def test_user_fork_starts_at_A_and_splits_at_M(self):
        layers,xy=layout(['Einspeisung','Verteiler','Behaelter','Behaelter'],[True,False,True,True],[(0,1),(1,2),(1,3)])
        self.assertEqual(layers,[0,1,2,2])
        self.assertLess(xy[0][0],xy[1][0]);self.assertLess(xy[1][0],xy[2][0])
        self.assertEqual(xy[2][0],xy[3][0]);self.assertNotEqual(xy[2][1],xy[3][1])
        self.assertEqual(xy[0][1],xy[1][1])
        self.assertEqual(xy[1][1],(xy[2][1]+xy[3][1])/2)
    def test_order_independent_source_choice_and_manual_positions(self):
        roles=['Behaelter','Verteiler','Einspeisung','Behaelter']
        layers,xy=layout(roles,[True,False,True,True],[(2,1),(1,0),(1,3)],manual={0:(17,23)})
        self.assertEqual(layers,[2,1,0,2]);self.assertEqual(xy[0],(17,23))
    def test_renderer_preserves_explicit_endpoints_and_labels(self):
        src=(ROOT/'vba/network/NetzDiagramm.bas').read_text()
        self.assertIn('msoConnectorStraight',src)
        self.assertNotIn('.RerouteConnections',src)
        self.assertIn('k(a).ID & " -> " & k(b).ID',src)
        self.assertIn('queue(head)',src)
        self.assertIn('k(i).Typ = "Einspeisung"',src)
    def test_user_network_hydraulics_match_analytic_fork(self):
        import math
        from tools.network.model import Node,Edge
        from tools.network.newton import solve
        nodes=[Node('A',3),Node('M'),Node('B',2),Node('S',2)]
        edges=[Edge('L1',0,1),Edge('L2',1,2),Edge('L3',1,3,d=.025)]
        r=solve(nodes,edges,mu=.1)
        resistance=128*.1*10/(math.pi*.05**4*3600*1e5)
        self.assertAlmostEqual(r['p'][1],2+16/33,7)
        for actual,numerator in zip(r['q'],[17,16,1]):
            self.assertAlmostEqual(actual,numerator/(33*resistance),5)
        self.assertAlmostEqual(r['q'][0],r['q'][1]+r['q'][2],6)
