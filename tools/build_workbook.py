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
    headers=['Q [m³/h]','v [m/s]','Re [-]','Darcy λ [-]','Rohr [bar]','Armaturen [bar]','Geodaetisch [bar]','Behaelter [bar]','Gesamt [bar]','Regime','Warnung']
    r.write_row('A4',headers,head);r.set_row(3,32);r.set_column('A:I',17,num);r.set_column('J:J',16);r.set_column('K:K',85)
    r.freeze_panes(4,1)
    r.conditional_format('J5:J10004',{'type':'text','criteria':'containing','value':'Uebergang','format':note})
    chart=w.add_chart({'type':'scatter','subtype':'straight'})
    for col,name in [(4,'Rohrreibung'),(5,'Einzelwiderstaende'),(8,'System gesamt')]:
        chart.add_series({'name':name,'categories':['Ergebnisse',4,0,104,0],'values':['Ergebnisse',4,col,104,col]})
    chart.set_title({'name':'Nach VBA-Berechnung automatisch aktualisiert'})
    chart.set_x_axis({'name':'Volumenstrom Q [m³/h]'});chart.set_y_axis({'name':'Differenzdruck Δp [bar]'})
    chart.set_legend({'position':'bottom'});chart.set_size({'width':900,'height':440});r.insert_chart('M4',chart)
    w.close()
    return path
if __name__ == '__main__':print(build())
