# Validierung und verbleibende Excel-Abnahme

## In der Cloud ausgeführt

`python -m unittest discover -s tests -v`: **15 Tests bestanden**.

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

- Module importieren, Projekt kompilieren, Einrichtung zweimal ausführen: genau vier eigene Schaltflächen.
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

Die Abnahme der Excel-Integration bleibt offen. Die erste Version ersetzt keine projektbezogene Prüfung der Stoffdaten, Widerstandsbeiwerte und Modellannahmen.
