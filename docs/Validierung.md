# Validierung und verbleibende Excel-Abnahme

## In der Cloud ausgeführt

`python -m unittest discover -s tests -v`: **38 Tests bestanden**.

Diese prüfen die XLSX-Paketintegrität, die drei Blätter, Tabellen und Namensbereiche, Eingabezellzuordnung, drei XY-Diagrammreihen, importierbare ASCII-Module und Referenzdaten. Numerisch prüfen sie Hagen-Poiseuille im laminaren Fall, einen bekannten glatten Colebrook-Wert, additive Einzelwiderstände, Druckvorzeichen, Q=0 und den Übergangsansatz. Eine Python-Transkription der VBA-Fixpunktiteration wird in **45 Kombinationen** (Re=2300 bis 10⁸, ε/D=0 bis 0,05) gegen eine unabhängige Bisektion direkt in λ geprüft.

Dies ist eine Prüfung des Algorithmus und der erzeugten Dateien. **VBA wurde hier weder kompiliert noch ausgeführt**, da Microsoft Excel/Windows nicht installiert ist. Das Python-Ergebnis belegt weder die Makroausführung noch die Bedienung in Excel.

## Automatisierte VBA-Selbsttests

`ValidierungStarten` ruft den tatsächlichen VBA-Rechenkern auf und vergleicht jeweils acht Ergebnisgrößen (v, Re, λ, Rohr-, Einzel-, Höhen-, Behälter- und Gesamtdruck) mit unabhängigen, vorab erzeugten Referenzwerten. Toleranz je Größe: 10⁻⁸·(1+|Referenz|), in der jeweiligen Ausgabeeinheit. Die Referenzberechnung verwendet Bisektion in λ und nicht die Fixpunktiteration des VBA-Solvers. Referenzdatei: `tests/reference_cases.json`.

Zwölf gültige Fälle:

1. Laminar: analytische Hagen-Poiseuille-Kontrolle.
2. Turbulent, hydraulisch glatt, ohne Armaturen.
3. Turbulent, rau, ohne Armaturen.
4. Mehrere Armaturen: 2×0,9 + 3×0,4 + 10 = 13.
5. Positive Höhendifferenz.
6. Q=0 mit Höhe und Behälterdruck.
7. Negativer statischer Gesamtdruck.
8. Übergangsbereich Re=3000.
9. Grenze Re≈2300.
10. Grenze Re≈4000.
11. Rohrlänge=0, nur Einzelwiderstände.
12. Maximal zugelassene relative Rauigkeit 0,05.

Zusätzlich acht Fehlerfälle: negatives Q, D=0, negative Länge, negative Rauigkeit, Dichte=0, Viskosität=0, negative ζ-Summe, ε/D>0,05. Gleitkommaarithmetik kann die exakte Regimezuordnung eines aus Q rekonstruierten Grenzwerts um Rundungsstellen beeinflussen; zusätzlich werden die Faktoren bei exaktem Re in Python geprüft.

Erwartete VBA-Meldung: **20 Rechenkern-Testfälle bestanden**. Bei Abweichungen erscheint eine Fehlermeldung und ein VBA-Laufzeitfehler; Fehler werden nicht als bestandene Prüfung ausgegeben. Diese 20 Fälle sind vorbereitet, **noch nicht in Excel ausgeführt**.

## Erforderliche Abnahme in Excel für Windows

- Module importieren, Projekt kompilieren, Einrichtung zweimal ausführen: genau sieben eigene Schaltflächen.
- VBA-Selbsttests starten; erwartete Meldung prüfen.
- Standardwerte berechnen: 101 Punkte; Q von 0 bis 30 m³/h, drei Kurven, Laufzeitpunkt, Q=0 ohne dynamischen Druckverlust.
- Auf 2 und dann 10000 Punkte umstellen; Endpunkte und Diagrammbereiche prüfen. Ungültige Werte 1, 2,5 und 10001 müssen abgelehnt werden.
- DN wechseln, Vorschlag übernehmen, D manuell ändern: Berechnung muss den manuellen D verwenden.
- Zwei 90°-Bögen und drei 45°-Bögen aktivieren: Σnζ=3. Einen Bogenwert überschreiben und wieder leeren; Datenbankwert muss wieder gelten.
- Neue Armaturenzeile hinzufügen, befüllen und wieder entfernen; Summen und Dropdown prüfen. Aktiven freien Widerstand ohne manuellen Wert ablehnen.
- Höhe=10 m, ρ=1000 kg/m³, Behälterdruck=1 bar, Qmin=0: erster Gesamtdruck muss 1,980665 bar ergeben.
- Fehler durch Einfügen erzwingen: leere/fehlerhafte Zelle, D=0, μ=0, negative Anzahl, gebrochene Anzahl, unbekannter Typ, Qmax≤Qmin, ε/D>0,05. Meldung und Ungültig-Markierung alter Ergebnisse prüfen.
- Laminar/Übergang/turbulent berechnen; Warnungen für Übergang und unzuverlässige konstante ζ im kleinen Re prüfen.
- Als XLSM speichern, schließen, erneut öffnen, erneut berechnen. Für den lokalen Betrieb darf keine Netzverbindung erforderlich sein.

## Zusätzliche Prüfungen Version 1.1

Die 12 zusätzlichen Cloud-Tests prüfen Blendenverluste über einen unabhängigen Ansatz aus Messdifferenzdruck und Druckrückgewinnungsfaktor (12 Kombinationen aus β und Cd), ungültige Geometrie, Umrechnung auf einen anderen lokalen Rohrdurchmesser, Pumpeninterpolation und Dichtebezug, ungültige/mehrdeutige Kennlinien, Extrapolationsverbot, einen analytisch lösbaren Schnittpunkt mit quadratischer Systemkurve, einen Schnittpunkt im zweiten Pumpensegment, Randlösungen, fehlende Schnittpunkte, Q=0 und die Ablehnung einer falschen Nullstelle bei einem Modellsprung.

`ErweiterungenValidieren` prüft den tatsächlichen VBA-Code mit **37 zusätzlichen Fällen**: 12 Blendenreferenzen, zwei Anlagenpunkte mit Blende (einschließlich Q=0), analytische Pumpenlösung und Residuum, stückweise lineare Kennlinie, beide Randpunkte und 10 Fehlerfälle einschließlich des realen Re=2300-Modellsprungs, ergänzt in Version 1.2 um fünf Kv-Referenzen und drei Kv-Fehlerfälle. Die Blendenreferenzen entstehen unabhängig über Druckrückgewinnung; die Pumpenreferenzen sind analytische Lösungen. Relative/absolute Misch-Toleranz 10⁻⁶·(1+|Referenz|); der eigentliche Pumpensolver verlangt ein Druckresiduum ≤10⁻⁸ bar.

`ValidierungStarten` startet beide Sätze: zuerst die 37 Erweiterungsfälle, dann die Erfolgsmeldung der 20 ursprünglichen Rechenkernfälle. Insgesamt **57 vorbereitete VBA-Testfälle**, weiterhin **nicht in Excel ausgeführt**. Die Cloud prüft Python-Transkriptionen und Dateistruktur, keine VBA-Laufzeit.

Zusätzliche erforderliche Excel-Abnahme:

- Einzelpunkt-Q in G3, einschließlich 0 und negativer Eingabe prüfen. Derselbe Q muss dieselben Werte wie der entsprechende Systemkurvenpunkt liefern, auch mit aktiver Blende.
- Blende aktivieren/deaktivieren, DN-Vorschlag übernehmen, Di überschreiben, d/Di=0,2 und 0,75 sowie unzulässige Werte prüfen. Der Blendenanteil L muss Teil von F bleiben und darf im Gesamtwert nicht doppelt vorkommen.
- Pumpen-Beispieldaten ersetzen, Zeile hinzufügen/entfernen und berechnen. Förderhöhe-Druck-Umrechnung mit zwei Dichten prüfen. N11−I11 muss O11 entsprechen und |O11| ≤10⁻⁸ bar sein.
- Q-Sortierung, doppelte Q, steigende/konstante H, negative H und leere Tabellenzeilen müssen abgelehnt werden. Fehlenden Schnittpunkt, Randlösung und nicht konvergierenden Modellsprung prüfen.
- Pumpendiagramm muss unabhängig vom zuvor berechneten Systemdiagramm aktuelle Daten und den berechneten Betriebspunkt zeigen.
- Alle acht Module in eine neue Vorlage importieren, kompilieren, speichern und nach erneutem Öffnen sämtliche Modi ausführen.

Die Abnahme der Excel-Integration bleibt offen. Die erste Version ersetzt keine projektbezogene Prüfung der Stoffdaten, Widerstandsbeiwerte und Modellannahmen.


## Ergänzungen Version 1.2

Sieben weitere Cloud-Tests: Kv-Referenzdruck für Wasser und andere Dichte, Serie und Q=0; unveränderter Druck bei Umrechnung auf unterschiedliche Hauptrohrdurchmesser; unabhängige lokale Formelreferenzen im Standalone-Blatt; plausibles Verhältnis Mess-/bleibender Druck; A4-Druckbereiche; frische Eingaben und eigene Diagrammdaten im Systembericht; Kv-Eingabe und Ausgabespalten.

Acht zusätzliche tatsächliche VBA-Prüfungen sind vorbereitet: fünf Kv-Referenzen einschließlich Integration in den Anlagenpunkt sowie drei unzulässige Kv-Eingaben. Insgesamt 37 Erweiterungs- plus 20 ursprüngliche Kernfälle. Die 34 Cloud-Tests führen weiterhin kein VBA und keine Excel-Formelauswertung aus.

Noch in Microsoft Excel zu prüfen:

- Alle **acht** Module in die neue Vorlage importieren, kompilieren und `Einrichtung` ausführen; Systembericht-Schaltfläche darf außerhalb des Druckbereichs erscheinen.
- Kv=10, Q=10 m³/h, rho=1000: eine Armatur erzeugt 1 bar; drei in Serie 3 bar. Kv deaktivieren; Kernwerte müssen wiederhergestellt sein. Ungültige Kv/Anzahl ablehnen.
- `Blende_Standalone` allein in eine neue makrofreie Arbeitsmappe kopieren. Eingaben ändern: Ergebnisse müssen neu berechnet werden. DN-Vorschlag, Q=0, ungültige Geometrie, Cd und Viskosität testen. Ungültige Eingaben dürfen keine verwendbaren Ergebniszahlen liefern.
- Eigenständige Blende gegen dieselben Parameter der Systemblende bei Rohrlänge 0 und ohne sonstige Verluste vergleichen. Messdruck und bleibenden Druck nicht verwechseln.
- Druckvorschau Standalone: eine A4-Seite; sämtliche Eingaben, Formeln, Ergebnisse und Grenzen lesbar, keine abgeschnittenen Zeilen.
- Systembericht ohne Pumpe und mit Pumpe erstellen. Werte mit den anderen Modi vergleichen. Nach Eingabeänderung Bericht neu erstellen; Chart muss aktuelle Daten und bei Pumpe den aktuellen Schnittpunkt zeigen. Bericht erneut ohne Pumpe erstellen: alte Pumpenserie muss verschwinden.
- Druckvorschau Bericht: Deckblatt auf Seite 1, vollständiger Anhang ab Seite 2, kein Steuerelement im Druckbereich. Metadaten beim Neuerstellen erhalten. Bei Berechnungsfehler muss die Ungültig-Markierung im Bericht sichtbar sein.

Die Datei-/Druckeinstellungen sind in der Cloud geprüft; Excel-Formelberechnung, VBA-Ausführung und tatsächliches Druckrendering bleiben offen.


## Eigenständige Kv-Berechnung

Vier zusätzliche Cloud-Tests prüfen die eigenständige Datei und deren Übereinstimmung mit dem Blatt in der Gesamtmappe, lokale Formelreferenzen, Eingabeprüfungen und A4-Druckeinstellungen. Die tatsächlichen arithmetischen Ergebnisformeln werden aus der XLSX gelesen und über einen eingeschränkten Python-AST-Auswerter gegen Referenzwerte für Wasser, abweichende Dichte, mehrere Armaturen und Q=0 geprüft. Das ist keine Ausführung der Excel-Funktionen IF/IFERROR/ISNUMBER und ersetzt keinen Excel-Test.

Zusätzliche Windows-Excel-Abnahme: `Kv_Standalone.xlsx` ohne Makros öffnen, Q/Kv/Dichte/Anzahl ändern, Ergebnis mit direkter Kv-Referenz vergleichen. Ungültige Zahlen, Kv=0, negative Dichte, nicht ganzzahlige Anzahl sowie Q=0 prüfen. Gleiche Werte im integrierten Blatt müssen gleiche Ergebnisse liefern. Druckvorschau auf eine lesbare A4-Seite mit Formeln, Ergebnissen, Annahmen und Metadaten prüfen. Das Blendenblatt und die separate Blenden-XLSX bleiben unverändert verfügbar.
