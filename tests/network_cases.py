"""Fictional cases shared with generated actual-VBA regression fixtures."""
import math
from tools.network.model import Node as N,Edge as E,G
from reference import orifice_reference

def cases():
    out={}
    out['single']=([N('A',3),N('B',2.98)],[E('L1',0,1)])
    out['series']=([N('A',3),N('M'),N('B',2.98)],[E('L1',0,1),E('L2',1,2)])
    for name,count in [('parallel_identical',2),('parallel_different',2),('parallel_three',3)]:
        out[name]=([N('A',3),N('B',2.98)],[E(f'L{i}',0,1,length=10 if name=='parallel_identical' else 10*(i+1)) for i in range(count)])
    out['collector']=([N('A',3),N('V'),N('X'),N('Y'),N('S'),N('B',2.98)],[E('in',0,1),E('x1',1,2),E('y1',1,3,length=20),E('x2',2,4),E('y2',3,4),E('out',4,5)])
    out['consumers']=([N('A',3),N('V'),N('B',2.98),N('C',external=-.1)],[E('in',0,1),E('one',1,2),E('two',1,3)])
    out['loop']=([N('A',3),N('X'),N('Y'),N('B',2.98)],[E('ax',0,1),E('ay',0,2,length=20),E('xb',1,3,length=30),E('yb',2,3),E('xy',1,2)])
    out['orifice']=([N('A',3),N('B',2.98)],[E('L',0,1,zorifice=orifice_reference(.05,.025,.61))])
    kv=10.;d=.05
    zkv=200*(3600*math.pi*d*d/4/kv)**2
    out['kv']=([N('A',3),N('B',2.98)],[E('L',0,1,zkv=zkv)])
    out['zero']=([N('A',3),N('B',3)],[E('L',0,1)])
    out['reverse']=([N('A',2.98),N('B',3)],[E('L',0,1)])
    out['elevation']=([N('A',3,height=0),N('B',3-1000*G*10/1e5,height=10)],[E('L',0,1)])
    out['mesh20x30']=([N(str(i),3 if i==0 else 2.8 if i==19 else None) for i in range(20)], [E(f'chain{i}',i,i+1) for i in range(19)]+[E(f'cross{i}',i,i+2,length=15+i) for i in range(11)])
    out['check_closed']=([N('A',2.98),N('B',3)],[E('L',0,1,zeta=2,check=True)])
    out['check_open']=([N('A',3),N('B',2.98)],[E('L',0,1,zeta=2,check=True)])
    out['inactive']=([N('A',3),N('B',2.98)],[E('on',0,1),E('off',0,1,active=False)])
    out['turbulent']=([N('A',3),N('B',2)],[E('L',0,1,length=100,rough=.000045)])
    return out
