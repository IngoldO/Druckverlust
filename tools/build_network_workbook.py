"""Add network inputs to a NEW V2 workbook. Never rewrite existing V1 artifacts.
Reuses the unchanged V1 generator inside a temporary output directory.
"""
from pathlib import Path
import shutil,tempfile
import xlsxwriter
import build_workbook as legacy
ROOT=Path(__file__).resolve().parents[1]

def add_network(w):
    head=w.add_format({'bold':True,'bg_color':'#185469','font_color':'white','text_wrap':True})
    title=w.add_format({'bold':True,'font_size':17,'font_color':'#185469'})
    entry=w.add_format({'bg_color':'#E4F3FC','num_format':'0.000000'})
    note=w.add_format({'text_wrap':True,'bg_color':'#FFF3CC'})
    cfg=w.add_worksheet('Netzwerk_Setup');cfg.set_column('A:A',48);cfg.set_column('B:C',22)
    cfg.write('A1','Modus B | Netzwerk, Phasen 2 und 3',title)
    cfg.merge_range('A2:F3','Eigene Stoffwerte und Eingaben, unabhaengig vom Einzelrohrmodus. VBA-Module aus vba/network zusaetzlich importieren. NetzwerkBerechnen / NetzwerkValidieren ueber Alt+F8. Diagramm und Netzwerkbericht folgen in spaeteren Phasen.',note)
    for i,(key,value,unit) in enumerate([('Dichte',1000,'kg/m³'),('Dynamische Viskositaet',100,'mPa·s'),('Iterationslimit',100,'—'),('Volumenbilanz-Toleranz',1e-6,'m³/h'),('Energie-Toleranz',1e-8,'bar')],4):
        cfg.write(i,0,key);cfg.write(i,1,value,entry);cfg.write(i,2,unit)
    cfg.merge_range('A12:F15','Fiktives laminares Beispiel. Externer Zufluss positiv, Entnahme negativ. Fester Absolutdruck und externe Flussvorgabe am selben Knoten werden abgelehnt. Druckknoten liefern den berechneten Reservoir-Zu-/Abfluss. Offener Auslass: lokalen Atmosphaerendruck explizit eingeben.',note)
    def table_sheet(name,table,headers,rows):
        ws=w.add_worksheet(name);ws.set_column(0,len(headers)-1,20)
        ws.write('A1',name,title);ws.write('A2','Zeilen als Excel-Tabelle erweiterbar. Keine Luecken/duplizierten IDs. Eingaben siehe docs/v2/Netzwerk.md.',note)
        ws.add_table(3,0,3+len(rows),len(headers)-1,{'name':table,'columns':[{'header':h,'format':entry} for h in headers],'data':rows,'style':'Table Style Medium 2'})
        ws.freeze_panes(4,1);return ws
    nodes=table_sheet('Netzwerk_Knoten','tblNetzKnoten',['ID','Bezeichnung','Typ','Hoehe m','Druck bar(a)','Extern Q m3h','X optional','Y optional'],[['A','Quelle','Einspeisung',0,3,None,None,None],['M','Verteilung','Verteiler',0,None,None,None,None],['B','Ziel','Behaelter',0,2.98,None,None,None]])
    nodes.data_validation('C5:C10004',{'validate':'list','source':['Einspeisung','Verteiler','Verbindungsknoten','Sammelkollektor','Verbraucher','Behaelter','Offener Auslass']})
    pipes=table_sheet('Netzwerk_Leitungen','tblNetzLeitungen',['ID','Bezeichnung','Start','Ende','DN','Di mm','Laenge m','Rauigkeit mm','Aktiv 0/1','Pumpe 0/1'],[['L1','Zulauf','A','M',50,50,10,0,1,0],['L2','Ablauf','M','B',50,50,10,0,1,0]])
    w.define_name('NetzKnotenIDs','=tblNetzKnoten[ID]');w.define_name('NetzLeitungsIDs','=tblNetzLeitungen[ID]')
    for col in ['C','D']:pipes.data_validation(f'{col}5:{col}10004',{'validate':'list','source':'=NetzKnotenIDs'})
    pipes.data_validation('E5:E10004',{'validate':'list','source':'=DNWerte'})
    fittings=table_sheet('Netzwerk_Armaturen','tblNetzArmaturen',['Leitung','Modell','Typ','Anzahl','Zeta manuell','Kv m3h','Lokal-Di mm','Loch-D mm','Cd','Aktiv 0/1'],[['L1','Zeta','Frei definierbarer Einzelwiderstand',1,0,None,None,None,None,0]])
    fittings.data_validation('A5:A10004',{'validate':'list','source':'=NetzLeitungsIDs'})
    fittings.data_validation('B5:B10004',{'validate':'list','source':['Zeta','Rueckschlag','Kv','Blende']})
    fittings.data_validation('C5:C10004',{'validate':'list','source':'=ArmaturenTypen'})
    table_sheet('Netzwerk_Pumpe','tblNetzPumpe',['Q m3h','H m'],[[0,40],[10,38],[20,32],[30,23],[40,12],[50,0]])
    out=w.add_worksheet('Netzwerk_Ergebnisse');out.set_column('A:R',19)
    out.write('A1','Netzwerkergebnisse',title)
    out.write('A2','NICHT BERECHNET | NetzwerkBerechnen starten',note)
    out.write('A3','Ergebnisse werden nur bei konvergierten und geprueften Bilanzen geschrieben.',note)

def build():
    base_class=xlsxwriter.Workbook
    class NetworkWorkbook(base_class):
        def close(self):
            if not self.fileclosed and Path(self.filename).name=='Druckverlust_Vorlage.xlsx':add_network(self)
            return super().close()
    oldroot=legacy.ROOT
    try:
        with tempfile.TemporaryDirectory(prefix='druckverlust-v2-') as td:
            legacy.ROOT=Path(td);xlsxwriter.Workbook=NetworkWorkbook
            source=legacy.build()
            target=ROOT/'excel/Druckverlust_Netzwerk_V2.xlsx'
            shutil.copyfile(source,target)
    finally:legacy.ROOT=oldroot;xlsxwriter.Workbook=base_class
    return target
if __name__=='__main__':print(build())
