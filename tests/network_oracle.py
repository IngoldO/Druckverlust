"""Independent reference: nested scalar edge inversions + nodal-pressure solve.
No mixed flow/pressure Newton iteration shared with production algorithm.
Colebrook uses bisection from the original independent reference.
"""
import numpy as np
from scipy.optimize import root,brentq
from tools.network.model import loss,G
from reference import factor

def oracle(nodes,edges,rho,mu):
    free=[i for i,n in enumerate(nodes) if n.pressure is None]
    fixed=[n.pressure for n in nodes if n.pressure is not None]
    def flows(p):
        q=[]
        for e in edges:
            if not e.active:q.append(0.);continue
            drive=p[e.start]-p[e.end]+rho*G*(nodes[e.start].height-nodes[e.end].height)/1e5
            if e.check and drive<=0:q.append(0.);continue
            if drive==0:q.append(0.);continue
            sign=np.sign(drive)
            def f(q):return loss(q,e,rho,mu,factor)-abs(drive)
            hi=1.
            while f(hi)<0:hi*=2
            q.append(sign*brentq(f,0,hi,xtol=1e-11))
        return np.array(q)
    def assemble(v):
        p=np.array([n.pressure or 0 for n in nodes],float);p[free]=v
        q=flows(p);b=np.array([n.external or 0 for n in nodes],float)
        for j,e in enumerate(edges):b[e.start]-=q[j];b[e.end]+=q[j]
        return p,q,b
    if free:
        result=root(lambda v:assemble(v)[2][free],np.full(len(free),np.mean(fixed)),tol=1e-10)
        if not result.success and max(abs(assemble(result.x)[2][free]))>1e-7:raise ValueError('reference failure')
        p,q,b=assemble(result.x)
    else:p,q,b=assemble([])
    return p,q
