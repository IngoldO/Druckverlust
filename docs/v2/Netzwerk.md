# V2 – Netzwerkberechnung, Schemagrafik, Kennlinie und Bericht

## Status

| Phase | Stand |
|---|---|
| 1 Bestandsaufnahme | Dokumentiert, Rückkehrstand und SHA-256-Inventur vorhanden |
| 2 Datenmodell | Knoten, aktive Leitungen, individuelle Armaturen, getrennte Stoffwerte und Netzpumpe implementiert |
| 3 Solver | Gedämpftes gemischtes Newton-Verfahren, separate Residuen, Null-/Rückströmung, Maschen, Rückschlaglogik, Fehlermeldungen implementiert; Cloud-Verfahrensprüfung bestanden |
| 4 Excel-Oberfläche | Tabellen, Dropdowns, Schaltflächen, Ergebnis- und Kennzahlenübersicht implementiert |
| 5 Netzwerkschema | Shapes mit Richtung, Druck, Q, IDs und Pumpensymbol; automatische Anordnung und explizit speicherbare X/Y-Positionen implementiert |
| 6 Pumpe/Systemkurve | Pumpe im gekoppelten Netz; Q-Sweep mit neuer Verteilung je Punkt und erforderlichem Einspeisedruck/Pumpen-dp implementiert |
| 7 Netzwerkbericht | Eigenständiger A4-Netzwerkbericht mit Schema, Eingabestand, Resultaten, Formeln, Annahmen und optionalem Kennliniendiagramm implementiert |
| 8 Abnahme | 70 Cloud-Tests bestanden; tatsächliche VBA-Ausführung, Migration einer externen V4-XLSM und Windows-Excel-/Druckabnahme offen |

Die bisherige Einzelrohr-Systemkurve bleibt vollständig vorhanden und ist von der neuen Netzwerk-Kennlinie unabhängig.

## Installation und bewusste Übernahme vorhandener Werte

1. Alte Arbeitsmappe unverändert aufbewahren. `excel/Druckverlust_Netzwerk_V2.xlsx` als **neue** XLSM speichern.
2. Alle acht bisherigen BAS-Dateien aus `vba/` und zusätzlich alle **zehn** BAS-Dateien aus `vba/network/` importieren. Keine neuen Module ersetzen alte Namen.
3. VBA-Projekt kompilieren. Bisheriges `Einrichtung` für Modus A ausführen. Danach `NetzwerkEinrichtung` ausführen: fünf Schaltflächen auf `Netzwerk_Setup` öffnen die Netzwerkfunktionen. Alternativ Alt+F8: `NetzwerkBerechnen`, `NetzwerkKennlinieBerechnen`, `NetzwerkBerichtErstellen`, `NetzwerkPositionenSpeichern`, `NetzwerkValidieren`.
4. `ValidierungStarten` muss weiterhin 37 Erweiterungs- und 20 Kernfälle melden. `NetzwerkValidieren` startet 25 Netzwerk-Kernfälle und 8 zusätzliche Funktionstests. Alle diese Makros sind vorbereitet, nicht hier in Excel ausgeführt.
5. Fiktive Werte in `Netzwerk_*` bearbeiten und `NetzwerkBerechnen` starten. Der Ergebnisstatus muss KONVERGIERT lauten; separate Residuen prüfen.
6. Alte Eingabewerte bei Bedarf bewusst in die entsprechenden bisherigen Blätter der neuen Datei übertragen. Keine automatische Migration, kein Überschreiben von Anwenderdateien. Für eine externe V4-XLSM muss zuerst deren konkreter Blatt-/Modulbestand geprüft werden.

Alternative für eine bestehende XLSM: eine Kopie verwenden; alle neun Netzwerkblätter aus der V2-Vorlage kopieren, anschließend ausschließlich alle zehn Netzwerkmodule importieren. Die bestehenden Tabellen `tblDatenbank` und der Name `DNWerte` müssen vorhanden sein. Die strukturierten Namen `NetzKnotenIDs` und `NetzLeitungsIDs` auf die kopierten Tabellen prüfen. Diese Variante ist ebenfalls erst nach Kompilieren und Regression in Windows-Excel freizugeben.

**Upgrade vom ersten V2-Stand:** Am einfachsten die neue V2-Vorlage separat verwenden und Eingabewerte bewusst übertragen. Bei einer Kopie der bisherigen V2-XLSM die drei zusätzlichen Blätter `Netzwerk_Diagramm`, `Netzwerk_Kennlinie`, `Netzwerk_Bericht` ergänzen, auf `Netzwerk_Setup` die Felder B17–B20 aus der Vorlage übernehmen und alle zehn Netzwerkmodule gezielt ersetzen/importieren. Alte acht V1-Module nicht ersetzen. Anschließend kompilieren, beide Testsätze ausführen und `NetzwerkEinrichtung` starten. Keine Anwenderdatei wird automatisch verändert.

Die makrofreien Standalone-Dateien werden weder verändert noch für Modus B benötigt. `Blende_Standalone` und `Kv_Standalone` bleiben auch in der V2-Gesamtvorlage enthalten.

## Tabellen und Einheiten

### Netzwerk_Setup

B5: Dichte kg/m³; B6: dynamische Viskosität mPa·s; B7: Iterationslimit (1–1000); B8: maximale freie Knotenbilanz m³/h; B9: maximale Energieverletzung bar. B17: Einspeiseknoten (bei aktiver Pumpe deren End-/Druckknoten); B18/B19: Qmin/Qmax in m³/h; B20: 2–1000 Kennlinienpunkte. Diese Werte sind von `Eingabe` unabhängig. Standard: 1000 kg/m³, 100 mPa·s, 100 Iterationen, 10⁻⁶ m³/h und 10⁻⁸ bar. Die hohe Viskosität gehört zum fiktiven laminaren Startbeispiel, nicht zu einer allgemeinen Stoffempfehlung.

### Netzwerk_Knoten / tblNetzKnoten

ID, Bezeichnung, Typ, Höhe m, optionaler Druck bar(a), optionales externes Q m³/h, optionale X/Y-Positionen in Excel-Punkten relativ zum Schemabereich. Beide Positionen gemeinsam leer (automatisch) oder gemeinsam >=0 eingeben. IDs sind nicht leer, Groß-/Kleinschreibung unterscheidet keine IDs. Die Typen Einspeisung, Verteiler, Verbindungsknoten, Sammelkollektor und Verbraucher beschreiben die Rolle; Randbedingungen werden explizit gesetzt. Behälter und offener Auslass benötigen einen positiven Absolutdruck. Lokaler Atmosphärendruck wird nicht stillschweigend angenommen.

Leere Druckzelle bedeutet unbekannter Druck. Leeres externes Q bedeutet kein äußerer Strom (0); positive Werte sind Einspeisungen, negative Werte Entnahmen. Eine explizite externe Flussvorgabe, auch 0, wird an einem Druckknoten abgelehnt: dort bestimmt der Solver den Reservoirstrom. **Berechnete Drücke werden niemals in Eingabespalten zurückgeschrieben.**

### Netzwerk_Leitungen / tblNetzLeitungen

ID, Bezeichnung, Start/Ende (Knoten-IDs über Dropdown), DN (Auswahlhilfe), tatsächlicher Di mm, Länge m, Rauigkeit mm, Aktiv 0/1 und Pumpe 0/1. Leeres Aktiv bedeutet aktiv; leeres Pumpenfeld bedeutet keine Pumpe. Positive Flüsse zeigen Start→Ende, negative Flüsse Ende→Start. Inaktive Leitungen werden aus der Topologie entfernt; ihre IDs bleiben eindeutig. Isolierte Knoten/Teilnetze werden abgelehnt.

DN überschreibt Di nicht. Es gibt keinen DN-Übernahmeknopf für Netzwerkleitungen; den tatsächlichen Di bewusst eintragen. Jede Leitung hat einen eigenen einheitlichen D. Automatische T-Stück-, Sammler-, Reduzier- oder Erweiterungsverluste werden nicht erfunden; projektspezifisch als Armaturen erfassen.

### Netzwerk_Armaturen / tblNetzArmaturen

Leitungs-ID, Modell, Typ, Anzahl, optionaler manueller ζ, Kv m³/h, lokaler Blenden-Di mm, Loch-D mm, Cd und Aktiv 0/1. Modelle:

- `Zeta`: vorhandene Armaturendatenbank oder explizite manuelle Überschreibung. Numerisches **0 bleibt gültig**. Freier Widerstand benötigt einen expliziten Wert.
- `Rueckschlag`: ζ wie oben; zusätzlich einseitig in Leitungsbezugsrichtung. Auch ein Datenbanktyp Rückschlagventil aktiviert diese Logik.
- `Kv`: ausschließlich der individuell angegebene Kv wird zu einem äquivalenten ζ umgerechnet.
- `Blende`: ausschließlich die individuelle Geometrie und Cd bestimmen den äquivalenten ζ für den bleibenden Verlust.

Pro Armatur genau ein Modell; ζ, Kv und Blende werden nicht doppelt addiert. Mehrere Zeilen können dieselbe Leitung adressieren. Anzahl ganzzahlig 0–100000; 0 deaktiviert den Beitrag. Die Startzeile ist explizit inaktiv. Leere Tabellenzeilen entfernen; aktive Armaturen auf inaktiven Leitungen werden abgelehnt. Modellgrenzen der bisherigen Datenbank, Kv- und Blendenmodelle bleiben bestehen. Alle Rückschlagventile einer Leitung sind in deren positiver Bezugsrichtung orientiert.

### Netzwerk_Pumpe / tblNetzPumpe

Eigene Q/H-Tabelle, unabhängig vom bisherigen Blatt `Pumpe`. Höchstens eine aktive Pumpenleitung; Architektur hält die Pumpenzuordnung am Element. Q beginnt bei 0 und steigt streng, H≥0 fällt streng. Keine Extrapolation und kein rückwärts betriebener Pumpenmodus. Die Leitung darf zusätzlich Rohr-/Armaturenverluste haben. Der Pumpenbetriebspunkt entsteht aus der gemeinsamen Lösung sämtlicher Knoten- und Leitungsgleichungen, nicht aus einer nachträglichen Einzelrohr-Schnittpunktsuche.

## Gleichungssystem

Eine konstante Dichte für das gesamte Netz. Q in m³/h, p in bar(a), z in m; μ intern Pa·s, D/ε intern m.

Freier Knoten i:

```text
Σ Q_ein − Σ Q_aus + Q_extern = 0
Massenresiduum [kg/s] = Volumenresiduum [m³/h] · ρ/3600
```

Aktive Leitung e von i nach j:

```text
p_i − p_j + ρg(z_i−z_j)/100000 + Δp_Pumpe(Q_e)
  − sgn(Q_e) · [Δp_R(|Q_e|)+Δp_ζ(|Q_e|)+Δp_Kv(|Q_e|)+Δp_Blende(|Q_e|)] = 0
```

Der bestehende `BerechnePunkt` erhält den Betrag von Q; die Netzwerkhülle setzt das Vorzeichen der irreversiblen Verluste. Geodätische Anteile erscheinen genau einmal zwischen Knoten; keine Addition des Einzelrohr-Behälterdrucks. Unterschiedliche Druckrandbedingungen ersetzen diesen Anteil im Netz. Das lokale Rohrmodell hat an beiden Enden desselben Elements denselben D; zusätzliche Geschwindigkeitssprünge/Mischverluste an Knoten benötigen geeignete Armaturenwerte.

Rückschlagventil mit `w=Verlust−Antrieb`:

```text
Q >= 0, w >= 0, Q*w = 0
min(Q/tolQ, w/tolP) = 0
```

Bei geschlossenem Ventil darf ein Rückdruck getragen werden; der ausgegebene Energiefehler ist die Verletzung dieser Ungleichung, nicht der physikalisch zulässige gehaltene Differenzdruck. Nicht eindeutig festgelegte freie Drücke hinter geschlossenen Ventilen werden zusätzlich abgelehnt. Geringfügige Q-Werte innerhalb der Volumentoleranz werden als nahe null gekennzeichnet.

## Numerik und Fehlerbehandlung

Gemischte Unbekannte: alle aktiven Leitungsströme plus alle freien Knotendrücke. Die Gleichungszahl passt dazu; ein Druckknoten liefert keine zusätzliche Kontinuitätsgleichung. Mindestens eine Druckreferenz und ein verbundenes Eingabenetz sind erforderlich. Mehrere Druckreservoirs sind zulässig und nicht pauschal überbestimmt.

Initiale freie Drücke aus mittlerem Reservoir-Gesamtpotenzial, an jeweilige Höhe angepasst; initiale Flüsse 0,1 der Flussskalierung, Pumpe bei halbem Kennlinienbereich. Variablen werden nach typischen Q-/p-Größen skaliert, Gleichungen nach separaten Toleranzen. Vorwärts differenzierte Jacobi-Matrix (Schritt 10⁻⁶·(1+|x|)); an Pumpengrenzen einseitige Differenz. Zeilenskalierte Gauß-Elimination mit partieller Pivotisierung; Pivotbetrag unter 10⁻¹² nach Skalierung erzeugt einen Fehler. Das Pivotverhältnis ist nur ein Diagnoseindikator, **keine berechnete Matrix-Konditionszahl**. Singuläre/ungünstige Systeme werden nicht mit willkürlichen Drücken vervollständigt.

Armijo-artige Dämpfung: Startschritt 1, höchstens 40 Halbierungen. Nur Schritte im Pumpenbereich werden ausgewertet. Norm der toleranzskalierten Residuen muss sinken; Abbruch nur, wenn alle geforderten Bilanzen erfüllt sind. Nach Konvergenz erfolgen eine erneute physikalische Bilanzprüfung und die Prüfung positiver Absolutdrücke. Widerstandslose passive Leitungen werden abgelehnt (Knoten gegebenenfalls zusammenlegen). Der bisherige unstetige Re=2300-Übergangsansatz wird nicht geändert; Fälle ohne echte Lösung an einem Modellsprung können verständlich als nicht konvergierend enden.

Bei Fehlern: `Erfolg=False`, Status mit Ursache und keine Ergebnisarrays. Der Excel-Adapter leert alte Netzwerkzahlen und schreibt UNGUELTIG. Kein stiller Rückfall auf zuletzt erfolgreiche Ergebnisse. Das dichte Verfahren hat O(n³) Aufwand; mindestens 20 Knoten/30 Leitungen sind als Entwicklungsfall geprüft, größere Netze benötigen Leistungsmessung in Excel.

## Ergebnisse und heutige Grenzen

`Netzwerk_Ergebnisse` enthält Solverstatus, Iterationen, Volumen-/Massenresiduum, Energieverletzung und Pivotindikator; pro Knoten p, Höhe und Bilanz; pro Leitung signiertes Q/v, Re, Darcy-λ, getrennte Rohr-/ζ-/Kv-/Blendenverluste, irreversiblen Gesamtverlust, Richtung und Energiefehler. Knoten mit festem Druck: ausgegebene Bilanz ist der Nettoüberschuss der angeschlossenen Leitungen; der erforderliche äußere Reservoirzufluss ist dessen **negativer** Wert. Sie ist kein verletztes Kontinuitätsresiduum.

Vorhandene Einzelrohr-, Pumpen- und Engineering-Funktionen sind unverändert. Stoff-/Modellgrenzen aus V1 gelten weiter; keine Kavitation-/Phasen-/NPSH-/Thermo-/Transientenprüfung. Konstante ζ-/Kv-/Cd-Werte bei kleinem Re sind besonders unsicher; die laminaren fiktiven Validierungsfälle prüfen die Gleichungen, nicht die Auslegungsgültigkeit solcher Kennwerte.


## Schemagrafik und manuelle Anordnung

`NetzwerkBerechnen` aktualisiert nach erfolgreicher Bilanzprüfung das Blatt `Netzwerk_Diagramm`. Symbole enthalten ID, Name, Rolle und Absolutdruck. Verteiler und Sammelkollektoren haben eigene Farben, Druckreservoirs eine stärkere Umrandung. Verbindungslinien erhalten die Leitungs-ID, signiertes Q und Pfeile in tatsächlicher Strömungsrichtung. Bei nahezu null Q werden keine Richtungspfeile angezeigt. Pumpen tragen ein grünes P-Symbol. Die Anordnung ist schematisch und nicht maßstäblich; bei komplexen Maschen/Parallelleitungen kann eine manuelle Nachbearbeitung nötig sein.

Knoten-Shapes auf dem Diagramm verschieben und **Knotenpositionen speichern** wählen. Dies schreibt X/Y ausdrücklich in die Knotentabelle. Die Tabelle darf zwischen Zeichnen und Speichern nicht umsortiert werden; die IDs werden kontrolliert. X/Y leeren stellt die automatische Anordnung wieder her. Nur eigene Shapes mit Präfix `net_` werden neu aufgebaut; bestehende Diagramme im Einzelrohrmodus und fremde Shapes bleiben erhalten. Bei einer fehlgeschlagenen Berechnung verschwindet das alte Schema und erhält den Status UNGUELTIG.

Der Diagrammdruckbereich folgt den erzeugten Shapes. A4 quer und Anpassung an eine Seite sind voreingestellt. Für große Netze ist ein größeres Papierformat oder der Druck über mehrere Seiten sinnvoll; die Druckvorschau ist vor Freigabe zu kontrollieren.

## Netzwerk-Systemkennlinie

Auf `Netzwerk_Setup` Einspeisung (B17), Qmin/Qmax (B18/B19) und Punktanzahl (B20) setzen. **Systemkennlinie berechnen** starten. Für jeden Q-Punkt wird das gesamte Netz neu gelöst; die Teilströme stehen auf `Netzwerk_Kennlinie` ab Spalte AA. Alle Q-Punkte müssen konvergieren; andernfalls werden Kurve und Tabelle verworfen und eine Alternative genannt. Die Eingabetabellen werden nicht verändert.

**Ohne aktive Pumpe:** Am ausgewählten Knoten wird nur für die Kennlinienrechnung die Druckvorgabe durch einen äußeren Zufluss Q ersetzt. Mindestens eine andere Druckreferenz muss erhalten bleiben. Weitere explizit positive äußere Einspeiseströme werden für diese Kurve abgelehnt; stattdessen den Netzwerkbetriebspunkt berechnen. Andere Entnahmen und Reservoirbedingungen bleiben unverändert. Ausgegeben wird der **absolute Portdruck bar(a)** und dessen Druckhöhe, nicht automatisch ein Pumpendifferenzdruck. Bei mehreren Druckreservoirs ist Q der kontrollierte Portstrom; andere Reservoirströme können sich ändern. Dies ist keine unbedingte Gesamtkennlinie für beliebig veränderte Randbedingungen.

**Mit aktiver Pumpe:** Als Einspeisung muss der End-/Druckknoten der Pumpenleitung gewählt werden. Der Solver erhält eine zusätzliche Unbekannte (erforderliche Pumpendruckerhöhung) und eine zusätzliche Gleichung `Q_Pumpe=Q_vorgegeben`. Die Herstellerkennlinie wird für diesen System-Q-Sweep nicht extrapoliert, sondern vorübergehend durch die gesuchte Druckerhöhung ersetzt. Die Topologie, Pumpenleitungsverluste und sämtliche anderen Randbedingungen bleiben erhalten. Ausgegeben werden erforderliches **Pumpen-dp in bar** und Förderhöhe in m. Die tatsächliche Herstellerkennlinie wird nur innerhalb ihrer Stützpunkte dargestellt. Der markierte Betriebspunkt wird getrennt durch die ursprüngliche gekoppelte Pumpen-/Netzrechnung bestimmt und ist nicht aus der grafischen Abtastung abgelesen.

Kein Druckanker, falscher Pumpenport, nicht eindeutige Drücke hinter gesperrten Rückschlagventilen oder fehlende Konvergenz führen zu einer verständlichen Meldung. Mit Rückschlagventil auf der Pumpenleitung ist Q=0 als erforderliche Förderhöhe nicht eindeutig und wird abgelehnt; Qmin>0 wählen oder den Netzwerkbetriebspunkt verwenden. Ein funktionierender Q-Sweep behauptet keinen vorhandenen Hersteller-Betriebspunkt: wenn dieser fehlt, wird es im Kurvenblatt ausdrücklich dokumentiert.

## Engineering-Bericht

Auf `Netzwerk_Bericht` Projekt, Dokumentnummer, Bearbeiter und Revision ausfüllen. **Engineering-Bericht erstellen** berechnet den Netzwerkbetriebspunkt aus den aktuellen Eingaben neu und baut den generierten Bereich ab Zeile 6 neu auf. Metadaten in Zeilen 3–4 bleiben erhalten. Alte Kurven oder Ergebnisse anderer Läufe werden nicht ungeprüft übernommen.

Der mehrseitige A4-Querformatbericht enthält Schema, Kennzahlen, Formeln, Annahmen, Freigabefeld, vollständige Knoten-/Leitungsdaten und Randbedingungen, getrennte Verlustanteile, Residuen, Armaturenparameter pro Leitung, wirksame summierte Beiwerte und Pumpenstützpunkte. Zusätzlich wird eine Kennlinie aus demselben Netzstand berechnet und mit Pumpe/Betriebspunkt dargestellt. Ist die gewählte Kennlinie nicht eindeutig oder ungültig, bleibt der gültige Betriebspunktbericht erhalten und nennt die fehlende Kennlinie mit Grund und Alternative; es wird keine scheinbar gültige Kurve gedruckt.

Diagrammdaten liegen im Berichtsblatt selbst. Nach Eingabeänderungen Bericht erneut erstellen. Ein fehlgeschlagener Netzwerkbetriebspunkt entfernt generierte Zahlen/Shapes und markiert den Bericht UNGUELTIG. Das Schema auf dem Deckbereich ist aus Platzgründen verkleinert; für große Netze das separate Diagramm vergrößert beilegen. Vor Engineering-Freigabe die tatsächliche Druckvorschau und Makroausführung in Microsoft Excel prüfen.
