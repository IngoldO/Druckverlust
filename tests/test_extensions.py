"""Algorithm transcriptions, not execution of VBA. Actual VBA fixtures are separate."""
import math
import unittest
from reference import orifice_reference, point


def orifice_impl(D,d,cd):
    if D <= 0 or d <= 0 or not 0 < cd < 1:raise ValueError
    b=d/D
    if not .2-1e-12<=b<=.75+1e-12:raise ValueError
    b=max(.2,min(.75,b))
    return (math.sqrt(1-b**4*(1-cd**2))/(cd*b*b)-1)**2


def pump(q,qs,hs,rho):
    if len(qs)<2 or len(qs)!=len(hs) or rho<=0:raise ValueError
    if any(x<0 for x in qs+hs):raise ValueError
    if any(b<=a for a,b in zip(qs,qs[1:])):raise ValueError
    if any(b>=a for a,b in zip(hs,hs[1:])):raise ValueError
    if not qs[0]<=q<=qs[-1]:raise ValueError
    for i in range(1,len(qs)):
        if q<=qs[i]:
            return (hs[i-1]+(hs[i]-hs[i-1])*(q-qs[i-1])/(qs[i]-qs[i-1]))*rho*9.80665/1e5


def solve(system,qs,hs,rho):
    lo,hi=qs[0],qs[-1]
    f=lambda q:pump(q,qs,hs,rho)-system(q)
    flo,fhi=f(lo),f(hi)
    if abs(flo)<=1e-8:return lo
    if abs(fhi)<=1e-8:return hi
    if flo<0 or fhi>0:raise ValueError('no bracket')
    for _ in range(120):
        mid=lo+(hi-lo)/2
        fm=f(mid)
        if abs(fm)<=1e-8:return mid
        if fm>0:lo=mid
        else:hi=mid
    raise ValueError('nonconverged/discontinuous')

class Extensions(unittest.TestCase):
    def test_orifice_against_pressure_recovery_reference(self):
        for b in [.2,.25,.5,.75]:
            for cd in [.55,.61,.7]:
                with self.subTest(beta=b,cd=cd):
                    self.assertAlmostEqual(orifice_impl(.05,.05*b,cd),orifice_reference(.05,.05*b,cd),10)
    def test_orifice_invalid_geometry(self):
        for args in [(0,.025,.61),(.05,0,.61),(.05,.05,.61),(.05,.001,.61),(.05,.025,0),(.05,.025,1)]:
            with self.subTest(args=args),self.assertRaises(ValueError):orifice_impl(*args)
    def test_local_diameter_conversion(self):
        q,rho=10,1000
        main,local,hole=.1,.05,.025
        zlocal=orifice_reference(local,hole,.61)
        zeq=zlocal*(main/local)**4
        vmain=q/3600/(math.pi*main**2/4)
        vlocal=q/3600/(math.pi*local**2/4)
        self.assertAlmostEqual(zeq*vmain**2,zlocal*vlocal**2,11)
    def test_pump_interpolation_and_density(self):
        self.assertAlmostEqual(pump(5,[0,10,20],[40,30,0],1000),3.4323275)
        self.assertAlmostEqual(pump(10,[0,10,20],[40,30,0],800),2.353596)
    def test_no_extrapolation(self):
        for q in [-.1,20.1]:
            with self.assertRaises(ValueError):pump(q,[0,20],[40,0],1000)
    def test_bad_pump_data(self):
        for qs,hs in [([0],[1]),([0,0],[2,1]),([1,0],[2,1]),([0,1],[1,2]),([0,1],[1,1]),([0,1],[-1,-2])]:
            with self.subTest(qs=qs,hs=hs),self.assertRaises(ValueError):pump(0,qs,hs,1000)
    def test_analytic_quadratic_system_intersection(self):
        rho,D,z=1000,.05,2
        k=z*rho/2/1e5/(3600*math.pi*D**2/4)**2
        expected=(-.1+math.sqrt(.01+16*k))/(2*k)
        qs,hs=[0,40],[4e5/(rho*9.80665),0]
        actual=solve(lambda q:k*q*q,qs,hs,rho)
        self.assertAlmostEqual(actual,expected,6)
    def test_piecewise_pump_intersection(self):
        self.assertAlmostEqual(solve(lambda q:1,[0,10,20],[30,20,0],1000),14.901418935,6)
    def test_endpoints(self):
        self.assertEqual(solve(lambda q:3.92266,[0,20],[40,0],1000),0)
        self.assertEqual(solve(lambda q:0,[0,20],[40,0],1000),20)
    def test_no_intersection(self):
        for system in [lambda q:5,lambda q:-1]:
            with self.assertRaises(ValueError):solve(system,[0,20],[40,0],1000)
    def test_discontinuity_does_not_become_false_root(self):
        with self.assertRaises(ValueError):solve(lambda q:0 if q<10 else 4,[0,20],[30,10],1000)
    def test_zero_flow_orifice_no_dynamic_loss(self):
        self.assertEqual(point(0,.05,10,0,1000,.001,0,orifice_reference(.05,.025,.61),0)[-1],0)
