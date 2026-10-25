import math
import unittest
import numpy as np
from tools.network.model import Node as N,Edge as E,G,loss
from tools.network.newton import solve
from network_oracle import oracle
from network_cases import cases

class Networks(unittest.TestCase):
    def test_independent_references_and_separate_residuals(self):
        for name,(nodes,edges) in cases().items():
            mu=.001 if name=='turbulent' else .1
            with self.subTest(case=name):
                r=solve(nodes,edges,mu=mu)
                p,q=oracle(nodes,edges,1000,mu)
                np.testing.assert_allclose(r['p'],p,atol=2e-7,rtol=1e-7)
                np.testing.assert_allclose(r['q'],q,atol=2e-5,rtol=1e-6)
                self.assertLessEqual(r['maxq'],1e-6);self.assertLessEqual(r['maxp'],1e-8)
    def test_laminar_analytic_single_series_parallel(self):
        all_cases=cases()
        resistance=128*.1*10/(math.pi*.05**4*3600*1e5)
        for name,multiplier in [('single',1),('series',.5),('parallel_identical',1)]:
            n,e=all_cases[name];r=solve(n,e,mu=.1)
            self.assertAlmostEqual(r['q'][0],.02/resistance*multiplier,6)
    def test_reverse_and_zero(self):
        for name in ['reverse','zero','elevation']:
            n,e=cases()[name];r=solve(n,e,mu=.1)
            if name=='reverse':self.assertLess(r['q'][0],0)
            else:self.assertLess(abs(r['q'][0]),1e-5)
    def test_pump_coupled_analytic(self):
        nodes=[N('A',2),N('B',2)]
        edges=[E('P',0,1,length=0,zeta=2,pump=True)]
        curve=([0,40],[4e5/(1000*G),0])
        result=solve(nodes,edges,curve=curve)
        k=2*1000/2/1e5/(3600*math.pi*.05**2/4)**2
        expected=(-.1+math.sqrt(.01+16*k))/(2*k)
        self.assertAlmostEqual(result['q'][0],expected,5)
    def test_missing_pressure(self):
        with self.assertRaisesRegex(ValueError,'pressure'):solve([N('A'),N('B')],[E('L',0,1)])
    def test_overdetermined(self):
        with self.assertRaises(ValueError):solve([N('A',3,external=0),N('B',2)],[E('L',0,1)])
    def test_disconnected(self):
        with self.assertRaisesRegex(ValueError,'disconnected'):solve([N('A',3),N('B',2),N('C')],[E('L',0,1)])
    def test_iteration_limit(self):
        n,e=cases()['single']
        with self.assertRaises(ValueError):solve(n,e,mu=.1,maxiter=1,tq=1e-14,tp=1e-14)
    def test_singular_ideal_edge(self):
        with self.assertRaises(ValueError):solve([N('A',3),N('B',2)],[E('L',0,1,length=0)])
    def test_no_pump_extrapolation(self):
        with self.assertRaises(ValueError):solve([N('A',2),N('B',200)],[E('P',0,1,length=0,pump=True)],curve=([0,10],[40,0]))
    def test_pressure_node_reaction_not_mass_residual(self):
        n,e=cases()['series'];r=solve(n,e,mu=.1)
        self.assertAlmostEqual(r['balance'][0],-r['q'][0],7)
        self.assertAlmostEqual(sum(r['balance']),0,7)
    def test_no_double_counting(self):
        e=E('L',0,1,length=0,zeta=2,zkv=3,zorifice=4)
        q=10.;v=q/3600/(math.pi*e.d**2/4)
        self.assertAlmostEqual(loss(q,e,1000,.001),9*1000*v*v/200000)

    def test_check_valve_cannot_leave_undetermined_pressure(self):
        # No arbitrary pressure may be emitted behind a shut check valve.
        with self.assertRaises(ValueError):
            solve([N('A',3),N('B',external=0)],[E('check',0,1,check=True)],mu=.1)

    def test_scaled_gauss_against_independent_lapack(self):
        from tools.network.newton import scaled_gauss
        rng=np.random.default_rng(20261025)
        for size in [2,10,50]:
            a=rng.normal(size=(size,size))+np.eye(size)*size
            a*=np.logspace(-8,8,size)[:,None]
            rhs=rng.normal(size=size)
            np.testing.assert_allclose(scaled_gauss(a,rhs),np.linalg.solve(a,rhs),rtol=1e-10,atol=1e-8)
        with self.assertRaises(ValueError):scaled_gauss(np.array([[1.,1.],[1.,1.]]),np.ones(2))
