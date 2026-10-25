# V2 – Validierung von Solver, Kennlinien und Engineering-Funktionen

## Ausgeführt

```bash
python -m pip install -r requirements-build.txt -r requirements-network-test.txt
python tools/build_network_workbook.py
python tools/build_network_validation.py
python tools/build_network_feature_tests.py
python tools/validate_network.py
python -m unittest discover -s tests -v
```

Alle Befehle im Repository-Verzeichnis ausführen. Installation ist nur für Entwicklungsprüfungen/Generierung nötig; Excel/VBA benötigt weder Python noch NumPy/SciPy.

Ergebnis: **67 Python-Tests bestanden**, davon alle 38 bisherigen Tests unverändert und 29 zusätzliche Tests. Die Regression prüft zusätzlich unveränderte SHA-256-Werte sämtlicher inventarisierter V1-Module, Generatoren, Tests und Arbeitsmappen. In der V2-Vorlage stimmen die Zellwerte/Formeln/Formatindizes der zehn alten Blätter sowie die bisherigen Tabellen- und Diagramm-XMLs mit V1 überein.

**Kein VBA wurde hier kompiliert oder ausgeführt.** Die bestehenden 20+37 VBA-Fälle sind bytegleich erhalten. `NetzwerkValidieren` ergänzt **25 Netzwerk-Kernfälle und 8 Funktionstests**, insgesamt damit 90 vorbereitete tatsächliche VBA-Testfälle zur späteren Windows-Abnahme. Python-Algorithmusprüfungen ersetzen diese Ausführung nicht.

## Unabhängige Referenz

`tools/network/newton.py` ist eine Python-Transkription des neuen gemischten Verfahrens zur Verfahrensprüfung. Gauß-Elimination, Schrittweiten, Dämpfung und Residuen entsprechen dem VBA-Entwurf. Die Gauß-Elimination wird zusätzlich gegen NumPy/LAPACK geprüft, einschließlich Zeilenskalierungen über 16 Größenordnungen und singulärer Matrix.

`tests/network_oracle.py` verwendet dagegen ein **anderes Lösungsverfahren**: nur freie Knotendrücke als Unbekannte, lokale skalare Inversion jeder Rohrkennlinie per Brent-Verfahren und einen SciPy-Knotenlöser. Colebrook wird dabei durch die bestehende unabhängige Bisektion in λ bestimmt. Für laminare Einzel-, Reihen- und Parallelfälle gibt es außerdem analytische Hagen-Poiseuille-Referenzen. Die Netzpumpe wird gegen einen analytisch lösbaren quadratischen Verlustansatz mit linear fallender Pumpenkennlinie geprüft, nicht gegen dieselbe Nullstellensuche.

Alle Daten sind fiktiv. Bei den meisten Topologietests ist μ=0,1 Pa·s gewählt, um analytisch überprüfbare laminare Fälle ohne Übergangssprünge zu erhalten. Ein eigener Test verwendet turbulente Strömung mit μ=0,001 Pa·s und Rauigkeit. Die Armaturenfälle prüfen die implementierte Konvention; sie belegen nicht die physikalische Gültigkeit konstanter ζ/Cd/Kv bei kleinem Re.

## Angeforderte Fälle

| Nr. | Fall | Prüfung |
|---|---|---|
| 1 | Einzelrohr | Analytisch und unabhängiges Verfahren |
| 2 | Zwei Rohre in Reihe | Analytisch und unabhängiges Verfahren |
| 3 | Zwei identische Parallelrohre | Analytisch und unabhängiges Verfahren |
| 4 | Zwei unterschiedliche Parallelrohre | Unabhängiges Verfahren, unterschiedliche L |
| 5 | Drei unterschiedliche Parallelrohre | Unabhängiges Verfahren |
| 6 | Gemeinsamer Sammelkollektor | Sechs Knoten, verzweigtes Netz |
| 7 | Unterschiedliche Verbraucher | Druckreservoir und separat vorgegebene Entnahme |
| 8 | Geschlossene Masche | Unsymmetrische Brücke mit Querverbindung |
| 9 | Drosselblende | Unabhängige ζ-Referenz, genau einmal addiert |
| 10 | Kv-Armatur | Kv→ζ-Umrechnung und Verlustvergleich |
| 11 | Nullströmung | Gleiche Potenziale, Q nahe null |
| 12 | Strömungsumkehr | Negatives Q bei vertauschter Druckdifferenz |
| 13 | Unterschiedliche Höhen | Hydrostatischer Gleichgewichtstest |
| 14 | Pumpenbetriebspunkt | Gemeinsame Netzgleichungen gegen analytische Lösung |
| 15 | Fehlender Druck | Ablehnung vor Solverstart |
| 16 | Überbestimmung | Ablehnung Druck plus äußerer Fluss am selben Knoten |
| 17 | Nicht verbundenes Teilnetz | Ablehnung isolierten Knotens |
| 18 | Nichtkonvergenz | Erzwungenes kleines Iterationslimit und enge Toleranz, kein Erfolg |

Zusätzlich: 20 Knoten/30 Leitungen mit Maschen; offene/geschlossene Rückschlagventile; undefinierter Druck hinter gesperrtem Ventil; inaktive Leitung; widerstandslose singuläre Verbindung; Pumpenbereich ohne zulässigen Betriebspunkt; Reservoirreaktionen; keine Doppelzählung. Diese Prüfungen laufen auf der Python-Transkription. 18 gültige passive Fälle liefern unabhängig berechnete Referenzdaten für VBA; die Pumpenreferenz ist analytisch. Fehlerszenarien und expliziter ζ=0 ergänzen das tatsächliche VBA-Testmakro auf 25 Fälle.

## Toleranzen und beobachtete Residuen

Standardkriterien: maximale Volumenbilanz an **freien** Knoten ≤10⁻⁶ m³/h; maximale Energiegleichungsverletzung ≤10⁻⁸ bar. Massenbilanz kg/s wird unabhängig als ρ/3600 mal Volumenbilanz ausgewiesen. Druckreservoirs liefern den offenen äußeren Austausch und werden nicht fälschlich gegen eine Nullbilanz geprüft.

Über die 18 gültigen passiven Referenznetze beobachtet:

| Größe | Größter beobachteter Wert |
|---|---:|
| Iterationen | 4 |
| Freie Knotenbilanz | 5,30·10⁻¹³ m³/h |
| Freie Massenbilanz | 1,48·10⁻¹³ kg/s |
| Energieverletzung | 7,64·10⁻⁹ bar |
| Druckabweichung zur unabhängigen Referenz | 4,09·10⁻¹² bar |
| Flussabweichung zur unabhängigen Referenz | 3,22·10⁻⁷ m³/h |

Genaue Werte pro Fall: `Validierungsergebnisse.json`, reproduzierbar mit `tools/validate_network.py`. Diese Zahlen beziehen sich ausschließlich auf den Python-Lauf. Sie sind weder eine garantierte Genauigkeit für beliebige Netze noch Excel-Laufzeitmessungen. Vergleichstoleranz VBA-Referenzen: 2·10⁻⁵ + 10⁻⁶·|Referenz|; zusätzlich werden im VBA-Makro die separat strengeren physikalischen Bilanzgrenzen geprüft.

## Noch erforderliche Windows-Excel-Abnahme

- V2 als neue XLSM speichern, acht alte und zehn Netzwerkmodule importieren und vollständig kompilieren.
- Alle bisherigen 57 VBA-Fälle ausführen; bestehende Einzelrohr-, Pumpen-, Berichts- und Standalone-Workflows gemäß V1-Anleitung prüfen.
- `NetzwerkValidieren` ausführen: 25 Kern- und 8 Funktionstests erwartet, einschließlich 20/30-Netz.
- Tabellen erweitern, Zeilen entfernen, Dropdowns und unbekannte/duplizierte IDs prüfen; befüllt berechnete Drücke niemals in Eingaben zurückschreiben.
- Aktive/inaktive Leitungen, ζ=0, individuelle Kv-/Blendenwerte, Strömungsumkehr und gesperrte Rückschlagventile mit Ergebnisresiduen vergleichen.
- Solverfehler provozieren: alte Netzwerkergebnisse müssen verschwinden und UNGUELTIG sichtbar sein. Modus-A-Eingaben bleiben erhalten.
- Gleiches Netz in beiden Modi aufsetzen und Einzelrohr-Vergleich durchführen; auf bar(a) gegenüber bisheriger Differenzdruckeingabe achten.
- Pumpenwerte ändern und Netzwerk neu lösen; niemals außerhalb der eingegebenen Q-Grenzen rechnen.

Schemagrafik, Netzwerk-Systemkennlinie und Netzwerkbericht sind inzwischen implementiert. Ihre Makroausführung und das Druckrendering benötigen weiterhin die nachfolgende Windows-Abnahme. Eine externe V4-XLSM benötigt eine eigene Bestandsaufnahme, sobald sie tatsächlich vorliegt.

## Zusätzliche Prüfungen für die vervollständigten Funktionen

Zehn zusätzliche Cloud-Tests prüfen analytischen Portdruck einer Reihenschaltung einschließlich Q=0, unveränderte Eingabedaten, Q-abhängige Verteilung in unterschiedlichen Parallelzweigen gegen das unabhängige Verfahren, erforderliche Pumpen-dp bei Q=0/10/50, eine geschlossene Umlaufschaltung mit Pumpe, Übereinstimmung von Systemkurve und gekoppeltem Herstellerbetriebspunkt, unzureichende Randbedingungen und eine nicht eindeutige Nullstrom-Pumpenkennlinie mit Rückschlagventil. Zusätzlich werden die neuen Arbeitsblätter, Setupfelder, A4-Querformat, Shape-Präfixe, Vorzeichenpfeile und frische Berichtsdaten strukturell geprüft.

Die tatsächlichen VBA-Funktionsprüfungen werden mit `tools/build_network_feature_tests.py` erzeugt. `NetzwerkFunktionenValidieren` enthält acht Fälle mit analytischen Referenzen und Fehlerprüfungen. `NetzwerkValidieren` startet diese zusätzlich zu den 25 bisherigen Netzwerkfällen. Alle alten V1-Module/-Prüfungen bleiben bytegleich. Die Python-Tests prüfen weiterhin kein Excel-Rendering und führen kein VBA aus.

Erweiterte Excel-Abnahme:

1. Neue V2-Vorlage als XLSM speichern; acht V1- plus zehn Netzwerkmodule importieren, kompilieren, `Einrichtung` und `NetzwerkEinrichtung` ausführen. Alle 90 vorbereiteten VBA-Fälle ausführen (20+37+25+8).
2. Modus A einschließlich Systembericht und beide Standalone-Blätter wie bisher prüfen.
3. Netzwerk rechnen: Knotendrücke, Teilströme, Richtungen, Pumpensymbol und farbige Verteiler/Sammler im Schema prüfen; bei negativem Q muss der Pfeil umkehren. Fremdes Shape anlegen, neu berechnen: es darf nicht gelöscht werden.
4. Knoten verschieben, Positionen explizit speichern, neu rechnen: Positionen müssen erhalten bleiben. X/Y gemeinsam leeren: automatische Anordnung. Knoten umsortieren: Speichern eines alten Schemas muss abgelehnt werden.
5. Q-Sweep am Einspeiseport: Bilanz pro Punkt prüfen; Verteilung in parallelen Leitungen darf neu berechnet werden. Originale Druck-/Flusseingaben dürfen nicht überschrieben werden.
6. Pumpe aktivieren, Druckknoten als Einspeisung wählen: Systemkurve muss erforderliche Pumpen-dp zeigen; Herstellerkurve nur im gültigen Bereich. Betriebspunkt mit direkter Netzwerkrechnung vergleichen. Falschen Port, fehlenden Druckanker und Q=0 mit Pumpen-Rückschlagventil kontrolliert ablehnen.
7. Projekt, Dokument, Bearbeiter und Revision im Bericht eintragen. Bericht zweimal erstellen, Eingaben dazwischen ändern: Metadaten erhalten, numerische Daten und Diagramme frisch. Druckvorschau: vollständiges Schema, lesbare mehrseitige A4-Querformat-Ausgabe und vollständiger Anhang. Große Schemen ggf. separat vergrößert beilegen.
8. Ungültige Kennlinie bei weiterhin gültigem Betriebspunkt erzeugen: Bericht muss die fehlende Kennlinie ausdrücklich begründen. Ungültigen Netzwerkbetriebspunkt erzeugen: keine alten Zahlen/Shapes als gültig stehen lassen.
9. Speichern, schließen, ohne Netzwerkverbindung erneut öffnen und alle Modi rechnen.

**Offener Rest:** vollständige VBA-Kompilierung/Ausführung, visuelle Prüfung und Druckabnahme in Microsoft Excel für Windows. Diese Prüfungen können in der vorhandenen Linux-Cloud nicht durchgeführt werden und sind nicht als bestanden ausgewiesen.
