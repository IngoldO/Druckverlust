"""Port/pump system characteristic; equivalent to NetzKennlinienPunkt API."""
from dataclasses import replace
from .newton import solve

def point(nodes,edges,feed,q,**options):
    if q<0 or not 0<=feed<len(nodes):raise ValueError('feed/q')
    pumps=[e for e in edges if e.active and e.pump]
    if pumps:
        if feed!=pumps[0].end:raise ValueError('select pump discharge node')
        if q==0 and pumps[0].check:raise ValueError('zero flow closed-check curve not unique')
        r=solve(nodes,edges,pump_flow=q,**options)
        return r['pump_pressure'],r
    copied=[replace(n) for n in nodes]
    copied[feed]=replace(nodes[feed],pressure=None,external=q)
    if not any(n.pressure is not None for n in copied):raise ValueError('missing outlet pressure; use operating point')
    if any((n.external or 0)>0 for i,n in enumerate(copied) if i!=feed):raise ValueError('other specified sources; use operating point')
    r=solve(copied,edges,**options)
    return r['p'][feed],r
