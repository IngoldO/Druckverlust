"""Rebuild the macro-free template. Build dependency: XlsxWriter 3.2.9."""
from pathlib import Path
import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
ARMATUREN = [
    ('90°-Rohrbogen', .9, 'Standardradius; stark von R/D abhaengig'),
    ('45°-Rohrbogen', .4, 'Standardradius; stark von R/D abhaengig'),
    ('T-Stück, gerader Durchgang', .6, 'Gleiche Durchmesser; Durchflussaufteilung beachten'),
    ('T-Stück, Abzweig', 1.8, '90 Grad Abzweig; Durchflussaufteilung beachten'),
    ('Kugelhahn, vollständig geöffnet', .05, 'Voller Durchgang, nur voll geoeffnet'),
    ('Absperrschieber, vollständig geöffnet', .15, 'Nur voll geoeffnet'),
    ('Sitzventil', 10., 'Konventionelle Bauform, voll geoeffnet'),
    ('Rückschlagventil', 2., 'Illustrativer Klappentyp; Herstellerdaten erforderlich'),
    ('Schmutzfänger', 2., 'Sauber; Siebgeometrie und Verschmutzung beachten'),
    ('Rohreintritt', .5, 'Scharfkantiger Eintritt aus grossem Behaelter'),
    ('Rohraustritt', 1., 'Austritt in grossen Behaelter, kinetische Energie dissipiert'),
    ('Frei definierbarer Einzelwiderstand', 0., 'Benutzer muss projektspezifischen Wert setzen'),
]
DNS = [(10,13.6),(15,17.3),(20,22.3),(25,28.5),(32,37.2),(40,43.1),(50,54.5),(65,70.3),(80,82.5),(100,107.1),(125,131.7),(150,159.3),(200,210.1),(250,263.0),(300,312.7),(350,344.4),(400,393.8),(500,493.0)]

def build():
    path = ROOT / 'excel' / 'Druckverlust_Vorlage.xlsx'
    path.parent.mkdir(exist_ok=True)
    w = xlsxwriter.Workbook(path)
    w.set_properties({'title':'Rohrleitungs-Systemkennlinie','comments':'VBA-Module importieren und als XLSM speichern; README beachten.'})
    title = w.add_format({'bold':True,'font_size':19,'font_color':'#185469'})
    head = w.add_format({'bold':True,'bg_color':'#185469','font_color':'white','text_wrap':True})
    entry = w.add_format({'bg_color':'#E4F3FC','num_format':'0.000'})
    integer_entry = w.add_format({'bg_color':'#E4F3FC','num_format':'0'})
    text_entry = w.add_format({'bg_color':'#E4F3FC'})
    note = w.add_format({'text_wrap':True,'font_color':'#6C5100','bg_color':'#FFF3CC'})
    num = w.add_format({'num_format':'0.000000'})
    e = w.add_worksheet('Eingabe'); r = w.add_worksheet('Ergebnisse'); db = w.add_worksheet('Armaturendatenbank')
    e.set_column('A:A',42); e.set_column('B:C',19); e.set_column('D:D',52); e.set_column('F:I',15)
    e.write('A1','Rohrleitungs-Systemkennlinie',title)
    e.merge_range('A2:D3','Blaue Zellen: Eingaben. Vor Verwendung VBA importieren, als XLSM speichern und Einrichtung ausfuehren. Nach jeder Aenderung neu berechnen. DN ist nur eine Auswahlhilfe.',note)
    params = [('Rohrnennweite DN',50,'—'),('Tatsaechlicher Innendurchmesser',54.5,'mm'),('Rohrlaenge',100,'m'),('Absolute Rauigkeit',.045,'mm'),('Hoehe Austritt minus Eintritt',0,'m'),('Dichte',998.2,'kg/m³'),('Dynamische Viskositaet',1.002,'mPa·s'),('Temperatur (optional)',20,'°C'),('Minimaler Volumenstrom',0,'m³/h'),('Maximaler Volumenstrom',30,'m³/h'),('Berechnungspunkte',101,'—'),('Druck Ziel minus Quelle',0,'bar')]
    for row,(label,value,unit) in enumerate(params,4):
        e.write(row,0,label); e.write(row,1,value,integer_entry if row in (4,14) else entry); e.write(row,2,unit)
    e.merge_range('F2:H2','Einzelpunkt Q [m³/h]',head)
    e.write('G3',10,entry)
    e.data_validation('G3',{'validate':'decimal','criteria':'>=','value':0})
    e.data_validation('B5',{'validate':'list','source':'=DNWerte'})
    for cell in ['B6','B10','B11']:
        e.data_validation(cell,{'validate':'decimal','criteria':'>','value':0})
    for cell in ['B7','B8','B13','B14']:
        e.data_validation(cell,{'validate':'decimal','criteria':'>=','value':0})
    e.data_validation('B15',{'validate':'integer','criteria':'between','minimum':2,'maximum':10000})
    e.write_comment('B6','DN-Vorschlag ist keine normative Rohrabmessung. Tatsaechlichen Innendurchmesser anhand Rohrdaten verifizieren.')
    e.merge_range('A18:D19','Zeta leer = aktueller Datenbankwert; Zahl = manuelle Ueberschreibung. Anzahl 0 deaktiviert die Zeile. Neue Zeilen per Schaltflaeche. Alle Standardwerte sind Naeherungswerte und auf die lokale Rohrgeschwindigkeit bezogen.',note)
    e.add_table('A20:D32',{'name':'tblArmaturen','style':'Table Style Medium 2','columns':[{'header':'Armaturentyp','format':text_entry},{'header':'Anzahl','format':integer_entry},{'header':'Zeta manuell','format':entry},{'header':'Hinweis'}],'data':[[a[0],0,None,'Leer = Datenbank; ueberschreibbar'] for a in ARMATUREN]})
    e.data_validation('A21:A32',{'validate':'list','source':'=ArmaturenTypen'})
    e.data_validation('B21:B32',{'validate':'integer','criteria':'between','minimum':0,'maximum':100000})
    e.data_validation('C21:C32',{'validate':'decimal','criteria':'>=','value':0,'ignore_blank':True})
    e.freeze_panes(4,1)
    db.set_column('A:A',42);db.set_column('B:B',14);db.set_column('C:D',65);db.set_column('F:G',16)
    db.write('A1','Armaturen und DN-Vorschlaege',title)
    db.merge_range('A2:D3','Illustrative Richtwerte, keine Herstellerkennwerte. Fuer Auslegung Bauform, Oeffnung, Reynolds-Zahl und Datenqualitaet pruefen. Quellen und Modellgrenzen siehe docs/Modell.md.',note)
    db.add_table('A4:D16',{'name':'tblDatenbank','style':'Table Style Medium 2','columns':[{'header':'Typ'},{'header':'Zeta'},{'header':'Bedingungen'},{'header':'Quelle / Datenqualitaet'}],'data':[[a,b,c,'Illustrativer Richtwert; Vergleich: Crane TP-410 / Idelchik, keine editionsgenaue Tabellenuebernahme'] for a,b,c in ARMATUREN]})
    db.write_row('F4',['DN','Di Vorschlag mm'],head)
    for row,item in enumerate(DNS,4):db.write_row(row,5,item)
    db.write_comment('G5','Illustrative Innendurchmesser einer Stahlrohrreihe; keine verbindliche Normzuordnung. Manuell verifizieren.')
    w.define_name('ArmaturenTypen',"='Armaturendatenbank'!$A$5:$A$16")
    w.define_name('DNWerte',"='Armaturendatenbank'!$F$5:$F$22")
    r.write('A1','Ergebnisse und Systemkennlinie',title)
    r.write('A2','NOCH NICHT BERECHNET – VBA importieren und Berechnung starten.',note)
    r.write('A3','Nach Aenderungen neu berechnen. Tabelleninhalt gilt nur fuer den letzten erfolgreichen Lauf.',note)
    headers=['Q [m³/h]','v [m/s]','Re [-]','Darcy λ [-]','Rohr [bar]','Armaturen [bar]','Geodaetisch [bar]','Behaelter [bar]','Gesamt [bar]','Regime','Warnung','davon Blende [bar]']
    r.write_row('A4',headers,head);r.set_row(3,32);r.set_column('A:I',17,num);r.set_column('J:J',16);r.set_column('K:K',85);r.set_column('L:L',22,num)
    r.freeze_panes(4,1)
    r.conditional_format('J5:J10004',{'type':'text','criteria':'containing','value':'Uebergang','format':note})
    chart=w.add_chart({'type':'scatter','subtype':'straight'})
    for col,name in [(4,'Rohrreibung'),(5,'Einzelwiderstaende'),(8,'System gesamt')]:
        chart.add_series({'name':name,'categories':['Ergebnisse',4,0,104,0],'values':['Ergebnisse',4,col,104,col]})
    chart.set_title({'name':'Nach VBA-Berechnung automatisch aktualisiert'})
    chart.set_x_axis({'name':'Volumenstrom Q [m³/h]'});chart.set_y_axis({'name':'Differenzdruck Δp [bar]'})
    chart.set_legend({'position':'bottom'});chart.set_size({'width':900,'height':440});r.insert_chart('N4',chart)
    b = w.add_worksheet('Lochblende')
    b.set_column('A:A',42); b.set_column('B:C',20); b.set_column('D:F',20)
    b.write('A1','Lochblende | bleibender Druckverlust',title)
    b.merge_range('A2:F3','Scharfkantige, duenne, konzentrische Blende. Cd ist eine Annahme; keine ISO-5167-Messauslegung. Rohr-Di und Lochdurchmesser manuell pruefen. Aktiv=0 deaktiviert die Blende.',note)
    bp=[('Aktiv (0=aus, 1=ein)',0,'—'),('Blenden-Nennweite DN',50,'—'),('Tatsaechlicher Rohr-Di an Blende',54.5,'mm'),('Lochdurchmesser d',25,'mm'),('Durchflussbeiwert Cd',.61,'—'),('Anzahl gleicher Blenden',1,'—')]
    for row,(label,value,unit) in enumerate(bp,4):
        b.write(row,0,label);b.write(row,1,value,integer_entry if row in (4,5,9) else entry);b.write(row,2,unit)
    b.data_validation('B5',{'validate':'integer','criteria':'between','minimum':0,'maximum':1})
    b.data_validation('B6',{'validate':'list','source':'=DNWerte'})
    for cell in ['B7','B8']:
        b.data_validation(cell,{'validate':'decimal','criteria':'>','value':0})
    b.data_validation('B9',{'validate':'decimal','criteria':'between','minimum':.000001,'maximum':.999999})
    b.data_validation('B10',{'validate':'integer','criteria':'between','minimum':1,'maximum':1000})
    b.merge_range('A12:F15','Modellbereich: 0,2 <= d/Di <= 0,75. Cd=0,61 ist ein illustrativer Startwert. Unter Re=10000 erscheint eine Unsicherheitswarnung. DN-Vorschlag ueber Schaltflaeche auf Eingabe uebernehmen; der tatsaechliche Di bleibt aenderbar. Bei abweichender Rohrweite sind Reduzierverluste separat zu ergaenzen.',note)
    b.write('A17','Durchmesserverhaeltnis beta')
    b.write_formula('B17','=IFERROR(B8/B7,"")',num,25/54.5)
    p = w.add_worksheet('Pumpe')
    p.set_column('A:C',23);p.set_column('E:L',13);p.set_column('N:O',22,num)
    p.write('A1','Pumpenkennlinie und Betriebspunkt',title)
    p.write('A3','NOCH NICHT BERECHNET',note)
    p.merge_range('A5:F8','BEISPIELWERTE ersetzen: Q streng steigend, Foerderhoehe H streng fallend. 2 bis 1000 Tabellenzeilen; keine Leerzeilen. H in Metern der gefoerderten Fluessigkeit. Berechnung ueber Pumpenbetriebspunkt auf Eingabe. Keine Extrapolation ausserhalb der eingegebenen Kennlinie.',note)
    p.merge_range('A10:F12','Pumpendruck in Spalte C wird beim Berechnen aus aktueller Dichte erzeugt. Nach jeder Aenderung erneut berechnen. Schnittpunkt und Verlustanteile stehen auf Betriebspunkte; das Diagramm erscheint rechts neben der Pumpentabelle.',note)
    p.add_table('A20:C26',{'name':'tblPumpe','style':'Table Style Medium 2','columns':[{'header':'Q [m3/h]','format':entry},{'header':'Foerderhoehe H [m]','format':entry},{'header':'Pumpendruck [bar]','format':num}],'data':[[q,h,None] for q,h in [(0,40),(10,38),(20,32),(30,23),(40,12),(50,0)]]})
    p.write_row('N4',['System Q [m3/h]','Systemdruck [bar]'],head)
    pump_chart=w.add_chart({'type':'scatter','subtype':'straight'})
    pump_chart.add_series({'name':'Pumpe','categories':['Pumpe',20,0,25,0],'values':['Pumpe',20,2,25,2]})
    pump_chart.add_series({'name':'System','categories':['Pumpe',4,13,204,13],'values':['Pumpe',4,14,204,14]})
    pump_chart.set_title({'name':'Nach Pumpenberechnung aktualisiert'})
    pump_chart.set_x_axis({'name':'Q [m³/h]'});pump_chart.set_y_axis({'name':'Differenzdruck [bar]'})
    pump_chart.set_size({'width':850,'height':440});p.insert_chart('E20',pump_chart)
    op=w.add_worksheet('Betriebspunkte')
    op.set_column('A:I',17,num);op.set_column('J:J',16);op.set_column('K:K',85);op.set_column('L:L',22,num);op.set_column('N:O',23,num)
    op.write('A1','Einzelpunkt und Pumpenbetriebspunkt',title)
    op.write('A2','Nach Aenderungen jeden gewuenschten Berechnungsmodus neu starten. Blendenanteil ist in Armaturen enthalten.',note)
    op.write('A3','Einzelpunkt: NOCH NICHT BERECHNET',note)
    op.write_row('A4',headers,head);op.set_row(3,32)
    op.write('A8','Schnittpunkt Pumpe / System',title)
    op.write('A9','Pumpenschnittpunkt: NOCH NICHT BERECHNET',note)
    op.write_row('A10',headers,head);op.set_row(9,32)
    op.conditional_format('J5:J11',{'type':'text','criteria':'containing','value':'Uebergang','format':note})
    w.close()
    return path
if __name__ == '__main__':print(build())
