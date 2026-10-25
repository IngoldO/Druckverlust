"""Print layouts: self-contained orifice calculation and engineering report."""

def print_layout(ws, end=44):
    ws.set_paper(9)  # A4
    ws.set_portrait()
    ws.set_margins(.3,.3,.35,.35)
    ws.fit_to_pages(1,1 if end==48 else 0)
    ws.print_area(f'A1:H{end}')
    ws.set_column('A:H',11)
    ws.set_default_row(16)
    ws.hide_gridlines(2)
    ws.set_header('&C&BDruckverlust | Engineering&B')
    ws.set_footer('&L&D&RSeite &P von &N')


def build_report_sheets(w, dns, include_system=True):
    title=w.add_format({'bold':True,'font_size':17,'font_color':'#185469','valign':'vcenter'})
    head=w.add_format({'bold':True,'bg_color':'#185469','font_color':'white'})
    text=w.add_format({'text_wrap':True,'valign':'vcenter','font_size':10})
    formula=w.add_format({'font_name':'Cambria','font_size':11,'bg_color':'#EDF3F5','text_wrap':True,'valign':'vcenter'})
    entry=w.add_format({'bg_color':'#E4F3FC','border':1,'num_format':'0.000','font_size':10})
    meta=w.add_format({'bg_color':'#E4F3FC','bottom':1,'font_size':10})
    number=w.add_format({'num_format':'0.000000','bold':True,'font_color':'#185469'})
    warn=w.add_format({'text_wrap':True,'bg_color':'#FFF3CC','font_size':10,'valign':'vcenter'})
    b=w.add_worksheet('Blende_Standalone'); print_layout(b,48)
    b.merge_range('A1:H2','Drosselblende | Berechnungsblatt',title)
    for label,cell,value in [('Projekt','A3',''),('Dokument','A4','')]:
        b.write(cell,label,text)
    b.merge_range('B3:D3','',meta);b.merge_range('B4:D4','',meta)
    b.write('F3','Bearbeiter',text);b.merge_range('G3:H3','',meta)
    b.write('F4','Revision',text);b.merge_range('G4:H4','',meta)
    b.merge_range('A6:H6','EINGABEN | lokale, unabhaengige Berechnung',head)
    params=[('Nennweite DN',50,'—'),('Tatsaechlicher Rohr-Di D',54.5,'mm'),('Blendendurchmesser d',25,'mm'),('Volumenstrom Q',10,'m³/h'),('Dichte rho',998.2,'kg/m³'),('Dyn. Viskositaet mu',1.002,'mPa·s'),('Durchflussbeiwert Cd',.61,'—'),('Temperatur (optional)',20,'°C')]
    for row,(label,val,unit) in enumerate(params,7):
        b.merge_range(row,0,row,2,label,text);b.merge_range(row,3,row,4,val,entry);b.write(row,5,unit,text)
    b.set_column('M:N',14,None,{'hidden':True})
    for row,item in enumerate(dns,4):b.write_row(row,12,item)
    b.data_validation('D8',{'validate':'list','source':'=$M$5:$M$22'})
    b.merge_range('G8:H8','Di-Vorschlag',text)
    b.merge_range('G9:H9','',number)
    b.write_formula('G9','=IFERROR(VLOOKUP(D8,$M$5:$N$22,2,FALSE),NA())',number,54.5)
    b.merge_range('G10:H15','DN legt Di nicht fest. D manuell anhand Rohrnorm und Wandstaerke pruefen. Cd=0,61 ist nur ein Startwert.',warn)
    b.merge_range('A17:C17','Eingaben gueltig (1=ja)',text)
    checks=['COUNTIF($M$5:$M$22,D8)=1','D9>0','D10>0','D11>=0','D12>0','D13>0','D14>0','D14<1','D10/D9>=0.2-1E-12','D10/D9<=0.75+1E-12','OR(ISBLANK(D15),AND(ISNUMBER(D15),D15>=-273.15))']
    checks += [f'ISNUMBER(D{i})' for i in range(8,15)]
    b.merge_range('D17:E17','',number)
    b.write_formula('D17','=IFERROR(IF(AND('+','.join(checks)+'),1,0),0)',number,0)
    b.merge_range('A18:H19','',warn)
    b.write_formula('A18','=IF(D17<>1,"UNGUELTIG: Eingaben pruefen; Ergebnisse nicht verwendbar.",IF(AND(D11>0,D22<10000),"WARNUNG: Re < 10000, konstantes Cd besonders unsicher.","Modellannahmen pruefen; keine Kavitationserkennung, keine ISO-5167-Konformitaet."))',warn,'Beim Oeffnen in Excel neu berechnen.')
    expressions={20:('Verhaeltnis beta','MAX(0.2,MIN(0.75,D10/D9))','—'),21:('Rohrgeschwindigkeit v','(D11/3600)/(PI()*(D9/1000)^2/4)','m/s'),22:('Rohr-Reynolds-Zahl','D12*D21*(D9/1000)/(D13/1000)','—'),23:('Bleibender Verlustbeiwert zeta','(SQRT(1-D20^4*(1-D14^2))/(D14*D20^2)-1)^2','—'),24:('Messdifferenzdruck','(1-D20^4)/(D14^2*D20^4)*D12*D21^2/200000','bar'),25:('Bleibender Druckverlust','D23*D12*D21^2/200000','bar'),26:('Geschwindigkeit im Loch','(D11/3600)/(PI()*(D10/1000)^2/4)','m/s')}
    for row,(label,expr,unit) in expressions.items():
        b.merge_range(row-1,0,row-1,2,label,text);b.merge_range(row-1,3,row-1,4,'',number)
        b.write_formula(row-1,3,f'=IF($D$17=1,{expr},NA())',number,'#N/A');b.write(row-1,5,unit,text)
    b.merge_range('A28:H28','BERECHNUNGSFORMELN | SI im Rechenkern',head)
    lines=['β = d/D; A = πD²/4; v = (Q/3600)/A; Re = ρvD/μ', 'a = √[1 − β⁴(1 − Cd²)]', 'ζbleibend = [a/(Cd β²) − 1]²', 'ΔpMess = [(1 − β⁴)/(Cd² β⁴)] ρv²/2', 'Δpbleibend = ζbleibend ρv²/2; Pa → bar: /100000', 'D,d: mm → m (/1000); μ: mPa·s → Pa·s (/1000)']
    for row,line in enumerate(lines,29):b.merge_range(row-1,0,row-1,7,line,formula);b.set_row(row-1,22)
    b.merge_range('A36:H36','ANNAHMEN UND GRENZEN',head)
    b.merge_range('A37:H41','Stationaere, einphasige inkompressible Newtonsche Fluessigkeit; konstante Stoffwerte. Duenne scharfkantige konzentrische Blende; 0,2 <= d/D <= 0,75; Cd konstant und vom Anwender zu bestaetigen. Messdifferenzdruck ist nicht bleibender Druckverlust. Druckrueckgewinnungsansatz im Zusammenhang mit ISO 5167-2; keine normgerechte Messstellenauslegung. Keine Pruefung von Kavitation, Flashing, Laerm, Erosion oder mechanischer Festigkeit. Temperatur ist nur dokumentiert.',text)
    b.merge_range('A43:H45','Blatt ohne VBA lauffaehig: Formeln werden in Excel automatisch neu berechnet. Alle Daten und die DN-Vorschlagstabelle liegen lokal auf diesem Blatt. Fuer eine Freigabe Stoffdaten, Cd und Modellgrenzen fachlich pruefen. Ausdruck/PDF: A4 Hochformat, eine Seite.',warn)
    b.merge_range('A47:H48','Geprueft durch: ____________________    Datum: ______________    Freigabe: ______________',text)
    if not include_system:
        return
    r=w.add_worksheet('Systembericht'); print_layout(r,44)
    r.set_print_scale(95)
    r.set_h_pagebreaks([44]);r.set_column('N:U',15,None,{'hidden':True})
    r.merge_range('A1:H2','Systemkennlinie | Engineering-Bericht',title)
    r.write('A3','Projekt',text);r.merge_range('B3:D3','',meta);r.write('F3','Bearbeiter',text);r.merge_range('G3:H3','',meta)
    r.write('A4','Dokument',text);r.merge_range('B4:D4','',meta);r.write('F4','Revision',text);r.merge_range('G4:H4','',meta)
    r.merge_range('A5:C5','Pumpe einschliessen (0/1)',text);r.write('D5',0,entry)
    r.data_validation('D5',{'validate':'integer','criteria':'between','minimum':0,'maximum':1})
    r.merge_range('A6:H7','NOCH NICHT BERECHNET | SystemberichtErstellen ausfuehren',warn)
    r.merge_range('A8:H10','Eingabestand wird beim Erstellen dokumentiert.',text)
    r.merge_range('A11:H13','Betriebspunkt / Hinweise werden beim Erstellen dokumentiert.',warn)
    r.merge_range('A30:H30','WICHTIGSTE GLEICHUNGEN',head)
    formulas=['v = (Q/3600)/(πD²/4); Re = ρvD/μ; laminar: λ = 64/Re', 'Colebrook: 1/√λ = −2 log10[ε/(3,7D) + 2,51/(Re√λ)]', 'ΔpR = λ(L/D)ρv²/2; ΔpE = Σ(nζ)ρv²/2', 'ΔpSystem = ΔpR + ΔpE + ρgΔh + ΔpBehaelter', 'Kv: Δp [bar] = n (ρ/1000) (Q/Kv)²; Blende: ζ = [√(1−β⁴(1−Cd²))/(Cdβ²)−1]²', 'Pumpe: ΔpP = ρgH(Q); Schnittpunkt: ΔpP − ΔpSystem = 0; Pa → bar: /100000']
    for row,line in enumerate(formulas,31):r.merge_range(row-1,0,row-1,7,line,formula);r.set_row(row-1,22)
    r.merge_range('A38:H41','Stationaer, inkompressibel, einphasig, Newtonsch; konstante Stoffwerte. Hauptrohr mit einheitlichem D. Re=2300–4000 unsicher (Maximum aus laminar/Colebrook). Zeta/Cd/Kv angenaehert, Herstellergrenzen pruefen. Kv: konstante Oeffnung, turbulent, nicht kavitierend. Pumpen-H streng fallend, linear interpoliert, keine Extrapolation. Keine NPSH-, Phasen- oder Festigkeitspruefung. Details und Eingabestand im Anhang.',text)
    r.merge_range('A43:H44','Geprueft durch: ____________________  Datum: _____________  Freigabe: ______________',text)
    r.merge_range('A46:H46','ANHANG | vollstaendiger Berechnungsstand',head)


def build_kv_standalone(w):
    """Local Excel formulas; can be copied alone, no VBA or other sheet needed."""
    title=w.add_format({'bold':True,'font_size':17,'font_color':'#185469','valign':'vcenter'})
    head=w.add_format({'bold':True,'bg_color':'#185469','font_color':'white'})
    text=w.add_format({'text_wrap':True,'valign':'vcenter','font_size':10})
    entry=w.add_format({'bg_color':'#E4F3FC','border':1,'num_format':'0.000'})
    meta=w.add_format({'bg_color':'#E4F3FC','bottom':1})
    number=w.add_format({'num_format':'0.000000','bold':True,'font_color':'#185469'})
    formula=w.add_format({'font_name':'Cambria','font_size':11,'bg_color':'#EDF3F5','text_wrap':True,'valign':'vcenter'})
    warn=w.add_format({'text_wrap':True,'bg_color':'#FFF3CC','font_size':10,'valign':'vcenter'})
    s=w.add_worksheet('Kv_Standalone');print_layout(s,48)
    s.merge_range('A1:H2','Kv-Armatur | Druckverlustberechnung',title)
    s.write('A3','Projekt',text);s.merge_range('B3:D3','',meta)
    s.write('A4','Dokument',text);s.merge_range('B4:D4','',meta)
    s.write('F3','Bearbeiter',text);s.merge_range('G3:H3','',meta)
    s.write('F4','Revision',text);s.merge_range('G4:H4','',meta)
    s.merge_range('A6:H6','EINGABEN | unabhaengig von der Systemberechnung',head)
    inputs=[('Volumenstrom Q',10,'m³/h'),('Kv bei aktueller Oeffnung',10,'m³/h'),('Dichte rho',1000,'kg/m³'),('Anzahl gleicher Armaturen in Serie',1,'—'),('Dyn. Viskositaet (dokumentiert)',1,'mPa·s'),('Temperatur (optional)',20,'°C')]
    for row,(label,value,unit) in enumerate(inputs,7):
        s.merge_range(row,0,row,2,label,text);s.merge_range(row,3,row,4,value,entry);s.write(row,5,unit,text)
    s.merge_range('G8:H13','Kv fuer die gewaehlte Oeffnung aus Herstellerdaten einsetzen. Kvs gilt nur bei voller Oeffnung. Anzahl gilt fuer gleiche Armaturen in Serie, nicht parallel.',warn)
    for cell in ['D9','D10','D12']:
        s.data_validation(cell,{'validate':'decimal','criteria':'>','value':0})
    s.data_validation('D8',{'validate':'decimal','criteria':'>=','value':0})
    s.data_validation('D11',{'validate':'integer','criteria':'between','minimum':1,'maximum':1000})
    s.merge_range('A16:C16','Eingaben gueltig (1=ja)',text);s.merge_range('D16:E16','',number)
    checks=[f'ISNUMBER(D{i})' for i in range(8,13)] + ['D8>=0','D9>0','D10>0','D11>=1','D11<=1000','D11=INT(D11)','D12>0','OR(ISBLANK(D13),AND(ISNUMBER(D13),D13>=-273.15))']
    s.write_formula('D16','=IFERROR(IF(AND('+','.join(checks)+'),1,0),0)',number,0)
    s.merge_range('A17:H19','',warn)
    s.write_formula('A17','=IF(D16<>1,"UNGUELTIG: Eingaben pruefen; Ergebnisse nicht verwendbar.","Kv-Naeherung: turbulent, konstante Oeffnung, nicht kavitierend. Herstellergrenzen pruefen; keine Viskositaetskorrektur.")',warn,'Beim Oeffnen in Excel neu berechnen.')
    results=[(21,'Druckverlust je Armatur','(D10/1000)*(D8/D9)^2','bar'),(22,'Gesamtdruckverlust in Serie','D11*(D10/1000)*(D8/D9)^2','bar'),(23,'Gesamtdruckverlust','D22*100000','Pa'),(24,'Verlusthoehe der Fluessigkeit','D22*100000/(D10*9.80665)','m')]
    for row,label,expr,unit in results:
        s.merge_range(row-1,0,row-1,2,label,text);s.merge_range(row-1,3,row-1,4,'',number)
        s.write_formula(row-1,3,f'=IF($D$16=1,{expr},NA())',number,'#N/A');s.write(row-1,5,unit,text)
    s.merge_range('A27:H27','BERECHNUNGSFORMELN | Kv-Definition',head)
    lines=['ρref = 1000 kg/m³; Kv: Volumenstrom bei Δp = 1 bar fuer Referenzwasser', 'Δpeinzeln [bar] = (ρ/ρref) · (Q/Kv)²', 'Δpgesamt [bar] = n · Δpeinzeln [bar]', 'Δp [Pa] = 100000 · Δp [bar]', 'hVerlust [m] = Δpgesamt [Pa] / (ρ g); g = 9,80665 m/s²', 'Q und Kv in m³/h; ρ in kg/m³; bei Q = 0 gilt Δp = 0']
    for row,line in enumerate(lines,28):s.merge_range(row-1,0,row-1,7,line,formula);s.set_row(row-1,22)
    s.merge_range('A35:H35','ANNAHMEN UND GRENZEN',head)
    s.merge_range('A36:H41','Stationaere, einphasige inkompressible Newtonsche Fluessigkeit; konstante Dichte und Armaturenoeffnung. Einfache Kv-Beziehung fuer turbulente, nicht kavitierende Stroemung. Keine Korrektur fuer Viskositaet, Reynolds-Zahl, FL, Kavitation, Flashing oder Choked Flow. Viskositaet und Temperatur sind nur dokumentiert. Kein Rohr-/Blendenmodell; DN und Lochdurchmesser gehoeren zur separaten Blendenberechnung. Kv-Definition im Kontext IEC 60534-2-1, keine vollstaendige normgerechte Regelventilauslegung.',text)
    s.merge_range('A43:H45','Ohne Makros und ohne Verknuepfungen zu anderen Blaettern. Ergebnisse aktualisieren sich mit Excel-Formeln. Fuer Engineering-Freigabe Kv, Stoffwerte und Anwendungsgrenzen pruefen. A4 Hochformat, eine Seite.',warn)
    s.merge_range('A47:H48','Geprueft durch: ____________________  Datum: ____________  Freigabe: ______________',text)
