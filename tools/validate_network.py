"""Reproduce algorithm/reference comparisons. Does not execute Excel or VBA."""
from pathlib import Path
import sys,json
import numpy as np
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT));sys.path.insert(0,str(ROOT/'tests'))
from network_cases import cases
from network_oracle import oracle
from tools.network.newton import solve

def report():
    rows=[]
    for name,(nodes,edges) in cases().items():
        mu=.001 if name=='turbulent' else .1
        r=solve(nodes,edges,mu=mu)
        p,q=oracle(nodes,edges,1000,mu)
        rows.append(dict(case=name,nodes=len(nodes),edges=len(edges),iterations=r['iterations'],max_volume_residual_m3h=r['maxq'],max_mass_residual_kgs=r['maxq']*1000/3600,max_energy_residual_bar=r['maxp'],max_reference_pressure_error_bar=float(np.max(abs(p-r['p']))),max_reference_flow_error_m3h=float(np.max(abs(q-r['q'])))))
    data={'scope':'Python algorithm transcription versus independent nested inversion/nodal reference; NOT VBA execution','cases':rows}
    path=ROOT/'docs/v2/Validierungsergebnisse.json'
    path.write_text(json.dumps(data,indent=2)+'\n')
    print(json.dumps({key:max(row[key] for row in rows) for key in ['iterations','max_volume_residual_m3h','max_mass_residual_kgs','max_energy_residual_bar','max_reference_pressure_error_bar','max_reference_flow_error_m3h']},indent=2))
if __name__=='__main__':report()
