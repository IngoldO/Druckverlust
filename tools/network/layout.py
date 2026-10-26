"""Layout algorithm transcription for geometry tests, not an Excel renderer."""
from collections import deque

def layout(roles,fixed,edges,manual=None,pump_start=None):
    root=next((i for i,t in enumerate(roles) if t=='Einspeisung'),None)
    if root is None:root=pump_start
    if root is None:root=next((i for i,p in enumerate(fixed) if p),0)
    layers=[-1]*len(roles);layers[root]=0;queue=deque([root])
    while queue:
        node=queue.popleft()
        for a,b in edges:
            other=b if a==node else a if b==node else None
            if other is not None and layers[other]<0:
                layers[other]=layers[node]+1;queue.append(other)
    if -1 in layers:raise ValueError('disconnected')
    counts={l:layers.count(l) for l in set(layers)}
    tallest=max(counts.values());slots={l:0 for l in counts};positions=[]
    for i,l in enumerate(layers):
        p=(l*300,(slots[l]+(tallest-counts[l])/2)*115);slots[l]+=1
        positions.append((manual or {}).get(i,p))
    return layers,positions
