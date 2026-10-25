"""Read-only baseline inventory. Run before V2 changes, at e8bd020."""
from pathlib import Path
import hashlib,json,re,subprocess,zipfile
import xml.etree.ElementTree as E
ROOT=Path(__file__).resolve().parents[1]
NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
def inventory():
    data={'commit':subprocess.check_output(['git','rev-parse','HEAD'],cwd=ROOT,text=True).strip(),'files':{},'modules':{},'workbooks':{},'python_tests':{}}
    for folder in ['vba','tools','tests','excel']:
        for p in sorted((ROOT/folder).glob('*')):
            if not p.is_file() or p.suffix not in ['.bas','.py','.xlsx','.xlsm']:continue
            rel=str(p.relative_to(ROOT));data['files'][rel]=hashlib.sha256(p.read_bytes()).hexdigest()
            if p.suffix=='.bas':data['modules'][rel]=re.findall(r'(?im)^Public (?:Sub|Function)\s+(\w+)',p.read_text())
            if p.suffix=='.py' and p.name.startswith('test_'):data['python_tests'][rel]=re.findall(r'def (test_\w+)',p.read_text())
            if p.suffix in ['.xlsx','.xlsm']:
                with zipfile.ZipFile(p) as z:
                    wb=E.fromstring(z.read('xl/workbook.xml'))
                    data['workbooks'][rel]={'sheets':[s.attrib['name'] for s in wb.findall('s:sheets/s:sheet',NS)],'names':[dict(n.attrib,text=n.text) for n in wb.findall('s:definedNames/s:definedName',NS)],'tables':[{**E.fromstring(z.read(n)).attrib} for n in z.namelist() if n.startswith('xl/tables/')],'charts':[n for n in z.namelist() if re.match(r'xl/charts/chart\d+\.xml$',n)],'embedded_vba':'xl/vbaProject.bin' in z.namelist()}
    return data
if __name__=='__main__':
    out=ROOT/'docs/v2/Bestand_V1.json'
    if out.exists():raise SystemExit('Baseline exists: do not overwrite.')
    out.write_text(json.dumps(inventory(),indent=2,ensure_ascii=False)+'\n')
