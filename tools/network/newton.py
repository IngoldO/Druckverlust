"""Python transcription of the mixed VBA Newton scheme for algorithm checks.
Not evidence of VBA compilation/execution. Independent oracle is in tests/network_oracle.py.
"""
import numpy as np
from .model import validate,loss,pump_dp,G

def scaled_gauss(jac,rhs):
    a=jac.copy();b=rhs.copy();n=len(b)
    scale=np.max(np.abs(a),axis=1)
    if np.any(scale==0):raise ValueError('singular')
    a/=scale[:,None];b/=scale
    for k in range(n):
        pivot=k+np.argmax(abs(a[k:,k]))
        if abs(a[pivot,k])<1e-12:raise ValueError('singular/ill-conditioned')
        a[[k,pivot],k:]=a[[pivot,k],k:];b[[k,pivot]]=b[[pivot,k]]
        for i in range(k+1,n):
            t=a[i,k]/a[k,k]
            a[i,k+1:]-=t*a[k,k+1:];b[i]-=t*b[k]
    dx=np.zeros(n)
    for i in range(n-1,-1,-1):dx[i]=(b[i]-a[i,i+1:]@dx[i+1:])/a[i,i]
    return dx

def solve(nodes,edges,rho=1000.,mu=.001,curve=None,maxiter=100,tq=1e-6,tp=1e-8):
    validate(nodes,edges,rho,mu)
    if maxiter<1 or tq<=0 or tp<=0:raise ValueError('limits')
    active=[(i,e) for i,e in enumerate(edges) if e.active]
    free=[i for i,n in enumerate(nodes) if n.pressure is None]
    m,nf=len(active),len(free)
    Qs=max([1]+[abs(n.external or 0) for n in nodes])
    Ps=max([1]+[n.pressure or 0 for n in nodes])
    mean=np.mean([n.pressure+rho*G*n.height/1e5 for n in nodes if n.pressure is not None])
    x=np.array([.1]*m+[(mean-rho*G*nodes[i].height/1e5)/Ps for i in free])
    pump=next((j for j,(_,e) in enumerate(active) if e.pump),None)
    if pump is not None:
        pump_dp(0,curve,rho);x[pump]=curve[0][-1]/2/Qs
    def residual(x):
        p=np.array([n.pressure or 0 for n in nodes],float);p[free]=x[m:]*Ps
        bal=np.array([n.external or 0 for n in nodes],float)
        eq=[]
        for j,(_,e) in enumerate(active):
            q=x[j]*Qs;bal[e.start]-=q;bal[e.end]+=q
            drive=p[e.start]-p[e.end]+rho*G*(nodes[e.start].height-nodes[e.end].height)/1e5
            if e.pump:drive+=pump_dp(q,curve,rho)
            r=(loss(q,e,rho,mu)-drive)/tp
            if e.check:r=min(q/tq,r)
            eq.append(r)
        return np.r_[bal[free]/tq,eq]
    for iteration in range(maxiter+1):
        f=residual(x);merit=np.max(np.abs(f))
        if merit<=1:break
        if iteration==maxiter:raise ValueError('iteration limit')
        jac=np.zeros((m+nf,m+nf))
        for j in range(m+nf):
            step=1e-6*(1+abs(x[j]))
            if pump==j and (x[j]+step)*Qs>curve[0][-1]:step=-step
            if pump==j and (x[j]+step)*Qs<0:step=(curve[0][-1]/Qs-x[j])/2
            xx=x.copy();xx[j]+=step
            jac[:,j]=(residual(xx)-f)/step
        dx=scaled_gauss(jac,-f)
        alpha=1.
        for _ in range(40):
            trial=x+alpha*dx
            if pump is None or 0<=trial[pump]*Qs<=curve[0][-1]:
                merit_trial=np.max(np.abs(residual(trial)))
                if merit_trial<merit*(1-1e-4*alpha) or merit_trial<=1:break
            alpha/=2
        else:raise ValueError('line search')
        x=trial
    p=np.array([n.pressure or 0 for n in nodes],float);p[free]=x[m:]*Ps
    if min(p)<=0:raise ValueError('nonphysical pressure')
    q=np.zeros(len(edges))
    for j,(i,e) in enumerate(active):q[i]=x[j]*Qs
    mass=np.array([n.external or 0 for n in nodes],float)
    energy=np.zeros(len(edges))
    for i,e in active:
        mass[e.start]-=q[i];mass[e.end]+=q[i]
        drive=p[e.start]-p[e.end]+rho*G*(nodes[e.start].height-nodes[e.end].height)/1e5
        if e.pump:drive+=pump_dp(q[i],curve,rho)
        energy[i]=loss(q[i],e,rho,mu)-drive
        if e.check and abs(q[i])<=tq:energy[i]=min(0,energy[i])
    connected={i for i,n in enumerate(nodes) if n.pressure is not None}
    for _ in nodes:
        for i,e in active:
            if e.check and abs(q[i])<=tq:continue
            if e.start in connected or e.end in connected:connected.update((e.start,e.end))
    if len(connected)!=len(nodes):raise ValueError('pressure isolated by closed checks')
    maxq=max([0]+[abs(mass[i]) for i in free]);maxp=max(abs(energy))
    if maxq>tq or maxp>tp:raise ValueError('residual postcheck')
    return dict(q=q,p=p,balance=mass,energy=energy,iterations=iteration,maxq=maxq,maxp=maxp)
