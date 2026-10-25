"""Independent Colebrook reference: bracketed bisection in Darcy lambda.
No iteration structure shared with VBA's fixed-point iteration in 1/sqrt(lambda).
"""
import math

def factor(re, rough):
    if re == 0: return 0.
    if re < 2300: return 64/re
    def residual(f): return 1/math.sqrt(f)+2*math.log10(rough/3.7+2.51/(re*math.sqrt(f)))
    lo,hi=.0001,1.
    assert residual(lo)>0 and residual(hi)<0
    for _ in range(100):
        mid=(lo+hi)/2
        if residual(mid)>0:lo=mid
        else:hi=mid
    f=(lo+hi)/2
    return max(f,64/re) if re<4000 else f

def point(q,d,l,eps,rho,mu,h,z,dp):
    v=q/3600/(math.pi*d*d/4)
    re=rho*v*d/mu
    f=factor(re,eps/d)
    dynamic=rho*v*v/2/1e5
    return [v,re,f,f*l/d*dynamic,z*dynamic,rho*9.80665*h/1e5,dp,f*l/d*dynamic+z*dynamic+rho*9.80665*h/1e5+dp]

CASES = [
 ('Laminar',(.1,.05,10,0,1000,.01,0,0,0)),
 ('Turbulent glatt',(10,.05,100,0,998.2,.001002,0,0,0)),
 ('Turbulent rau',(30,.05,100,.0001,1000,.001,0,0,0)),
 ('Mehrere Armaturen',(10,.05,100,.000045,1000,.001,0,2*.9+3*.4+10,0)),
 ('Geodaetisch',(10,.05,100,0,1000,.001,12,0,0)),
 ('Stillstand',(0,.05,100,.000045,1000,.001,12,3,1.5)),
 ('Negativer statischer Anteil',(0,.05,0,0,1000,.001,-10,0,-.5)),
 ('Uebergang',(3000*.001*math.pi*.05/4/1000*3600,.05,10,0,1000,.001,0,0,0)),
 ('Re2300',(2300*.001*math.pi*.05/4/1000*3600,.05,10,0,1000,.001,0,0,0)),
 ('Re4000',(4000*.001*math.pi*.05/4/1000*3600,.05,10,0,1000,.001,0,0,0)),
 ('Nur Einzelwiderstaende',(10,.05,0,0,1000,.001,0,5,0)),
 ('Maximale Rauigkeit',(10,.05,10,.0025,1000,.001,0,0,0)),
]


def orifice_reference(pipe_d, hole_d, cd):
    """Independent route: metering differential multiplied by recovery ratio."""
    if pipe_d <= 0 or hole_d <= 0 or not 0 < cd < 1:
        raise ValueError('orifice geometry/Cd')
    beta=hole_d/pipe_d
    if not .2-1e-12 <= beta <= .75+1e-12:raise ValueError('beta')
    beta=max(.2,min(.75,beta))
    tap_zeta=(1-beta**4)/(cd**2*beta**4)
    root=math.sqrt(1-beta**4*(1-cd**2))
    recovery=(root-cd*beta**2)/(root+cd*beta**2)
    return tap_zeta*recovery
