"""Development reference model; production runtime remains VBA only."""
from dataclasses import dataclass
import math

@dataclass
class Node:
    id: str
    pressure: float|None = None
    external: float|None = None
    height: float = 0.

@dataclass
class Edge:
    id: str
    start: int
    end: int
    d: float = .05
    length: float = 10.
    rough: float = 0.
    zeta: float = 0.
    zkv: float = 0.
    zorifice: float = 0.
    active: bool = True
    pump: bool = False
    check: bool = False

G=9.80665

def validate(nodes,edges,rho,mu):
    if len(nodes)<2 or rho<=0 or mu<=0:raise ValueError('nodes/properties')
    if len({n.id.casefold() for n in nodes})!=len(nodes) or any(not n.id.strip() for n in nodes):raise ValueError('node IDs')
    if len({e.id.casefold() for e in edges})!=len(edges) or any(not e.id.strip() for e in edges):raise ValueError('edge IDs')
    if not any(n.pressure is not None for n in nodes):raise ValueError('missing pressure')
    for n in nodes:
        if n.pressure is not None and (n.pressure<=0 or n.external is not None):raise ValueError('overdetermined or pressure')
    active=[e for e in edges if e.active]
    if sum(e.pump for e in active)>1:raise ValueError('one pump only')
    seen={0}
    for e in active:
        if e.start==e.end or not 0<=e.start<len(nodes) or not 0<=e.end<len(nodes):raise ValueError('endpoint')
        if e.d<=0 or e.length<0 or not 0<=e.rough/e.d<=.05 or min(e.zeta,e.zkv,e.zorifice)<0:raise ValueError('geometry')
        if e.length==0 and e.zeta+e.zkv+e.zorifice==0 and not e.pump:raise ValueError('ideal zero resistance edge')
    for _ in nodes:
        for e in active:
            if e.start in seen or e.end in seen:seen.update((e.start,e.end))
    if len(seen)!=len(nodes):raise ValueError('disconnected')

def fixed_factor(re,rough):
    if re==0:return 0.
    if re<2300:return 64/re
    x=7.
    for _ in range(200):
        xn=-2*math.log10(rough/3.7+2.51*x/re)
        if abs(xn-x)<1e-10:return max(1/xn**2,64/re) if re<4000 else 1/xn**2
        x=xn
    raise ValueError('Colebrook')

def loss(q,e,rho,mu,factor=fixed_factor):
    v=abs(q)/3600/(math.pi*e.d**2/4)
    re=rho*v*e.d/mu
    dp=(factor(re,e.rough/e.d)*e.length/e.d+e.zeta+e.zkv+e.zorifice)*rho*v*v/2/1e5
    return math.copysign(dp,q)

def pump_dp(q,curve,rho):
    qs,hs=curve
    if len(qs)<2 or len(qs)!=len(hs) or qs[0]!=0 or any(b<=a for a,b in zip(qs,qs[1:])) or any(b>=a for a,b in zip(hs,hs[1:])) or min(qs+hs)<0:raise ValueError('pump data')
    if q<0 or q>qs[-1]:raise ValueError('no extrapolation')
    for i in range(1,len(qs)):
        if q<=qs[i]:return rho*G/1e5*(hs[i-1]+(hs[i]-hs[i-1])*(q-qs[i-1])/(qs[i]-qs[i-1]))
