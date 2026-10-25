# Druckverlust – Excel/VBA-Systemkennlinie

Erste Version für inkompressible, einphasige Newtonsche Flüssigkeiten mit konstanten Stoffwerten und einem einheitlichen Rohrinnendurchmesser. Berechnet Darcy-Weisbach-Rohrreibung, Einzelwiderstände, geodätische Höhe und einen konstanten Behälterdifferenzdruck.

## Installation in Microsoft Excel für Windows

1. `excel/Druckverlust_Vorlage.xlsx` lokal öffnen und **Speichern unter → Excel-Arbeitsmappe mit Makros (*.xlsm)** wählen, beispielsweise `Druckverlust.xlsm`.
2. Mit **Alt+F11** den VBA-Editor öffnen. Im Projekt-Explorer das Projekt **dieser Arbeitsmappe** auswählen.
3. Über **Datei → Datei importieren** alle vier Dateien aus `vba/` importieren: `Hydraulik.bas`, `Eingaben.bas`, `Oberflaeche.bas`, `Validierung.bas`. Keine Module in PERSONAL.XLSB importieren. Vorhandene Module derselben Namen vor einer Aktualisierung gezielt ersetzen.
4. **Debuggen → VBAProject kompilieren** ausführen. Es sind keine zusätzlichen Verweise oder Add-ins erforderlich.
5. Zurück zu Excel; **Alt+F8 → Einrichtung → Ausführen**. Dies legt vier Schaltflächen auf `Eingabe` an. Wiederholtes Ausführen ersetzt nur die eigenen Schaltflächen.
6. Speichern. Makros nur für die geprüfte Arbeitsmappe gemäß der eigenen Unternehmensrichtlinie freigeben. Eine generelle Absenkung der Excel-Makrosicherheit ist nicht erforderlich.
7. **Validierung starten** ausführen. Erwartet: **20 Rechenkern-Testfälle bestanden**. Anschließend **Systemkennlinie berechnen** anklicken.

Eine fertige XLSM wird nicht mitgeliefert: In der Cloud steht kein Microsoft Excel zur Verfügung, um ein VBA-Projekt einzubetten, zu kompilieren und auszuführen. Die XLSX enthält absichtlich keine scheinbar bereits berechneten Resultate. Nach Import und Speichern läuft die Anwendung vollständig lokal ohne Python oder externe Dienste.

## Bedienung

- Blaue Zellen auf `Eingabe` bearbeiten. DN wählen und bei Bedarf **DN-Vorschlag übernehmen** drücken. Dies ersetzt B6 ausdrücklich; anschließend den tatsächlichen Innendurchmesser anhand Rohrnorm und Wandstärke überprüfen. Die Berechnung verwendet ausschließlich B6, nicht DN.
- Länge und Rauigkeit dürfen null sein; Durchmesser, Dichte und Viskosität müssen positiv sein. 0 ≤ Qmin < Qmax; 2–10000 ganzzahlige Berechnungspunkte. Positive Höhe bedeutet höher gelegenen Austritt. Positiver Behälterdruck bedeutet höherer Druck im Zielbehälter.
- In der Armaturentabelle den Typ wählen und eine ganzzahlige Anzahl eingeben. **Zeta manuell leer** verwendet den aktuellen Datenbankwert, eine Zahl einschließlich 0 überschreibt ihn. Nach Typwechsel bleibt eine manuelle Überschreibung bewusst erhalten; zum Verwenden des neuen Richtwerts die Zelle leeren. Aktive freie Widerstände benötigen eine explizite Eingabe.
- Weitere Zeilen mit **Armatur hinzufügen** anlegen. Anzahl 0 deaktiviert eine Komponente. Die Tabelle kann über die Excel-Tabellenfunktionen verkleinert werden; ihre Namen und die ersten drei Spalten nicht ändern.
- **Nach jeder Eingabe-, Tabellen- oder Datenbankänderung neu berechnen.** Es gibt keine automatisch laufenden Änderungsereignisse. Die Berechnung aktualisiert Ergebnistabelle und alle drei XY-Kurven in einem Lauf. `Ergebnisse!A2` zeigt den Laufzeitpunkt und die verwendete Summe n·ζ. Die Diagrammkurven sind für gerenderte Berechnungspunkte verbunden, keine Interpolation des Strömungsmodells.
- Warnungen stehen pro Punkt in Spalte K; der Übergangsbereich wird zusätzlich farblich markiert. Bei ungültigen Eingaben werden vorhandene Ergebnisse als **nicht aktuell** markiert, nicht als neuer Lauf ausgegeben.
- Datenbank-Richtwerte sind ausdrücklich Näherungen. Für eine technische Auslegung geeignete Herstellerwerte und Randbedingungen einsetzen. Werte beziehen sich auf die Geschwindigkeit im eingegebenen Rohrdurchmesser.

## Dateien und Entwicklung

```text
excel/Druckverlust_Vorlage.xlsx   Vorbereitete, makrofreie Arbeitsmappe
vba/Hydraulik.bas                 Rechenkern, SI-Einheiten
vba/Eingaben.bas                  Eingabeprüfung, Datenbank, DN und Tabellenzeilen
vba/Oberflaeche.bas               Ablaufsteuerung, Schaltflächen und Diagramm
vba/Validierung.bas               Automatisierte Excel-Rechenkernprüfungen
docs/Modell.md                   Gleichungen, Gültigkeit und Datenqualität
docs/Validierung.md              Prüfumfang, Ergebnisse und Excel-Abnahme
tests/reference.py               Unabhängige Colebrook-Referenz (Bisektion)
tests/reference_cases.json       Referenzdaten für VBA-Selbsttests
tests/test_validation.py         Cloud-Validierung und XLSX-Strukturprüfungen
tools/build_workbook.py          Reproduzierbarer Aufbau der XLSX-Vorlage
tools/build_validation.py        Erzeugung der VBA-Referenztests
```

Entwicklung benötigt Python 3.10+ und, ausschließlich zum Erzeugen der Vorlage, `XlsxWriter==3.2.9`:

```bash
python -m pip install -r requirements-build.txt
python tools/build_workbook.py
python tools/build_validation.py
python -m unittest discover -s tests -v
```

Diese Befehle aus dem Repository-Verzeichnis ausführen. Die beiden Generatoren überschreiben ausschließlich ihre eigenen generierten Artefakte: XLSX, Referenz-JSON und `Validierung.bas`. Änderungen an einer lokal angepassten Arbeitsmappe vorher separat sichern. Die importierbaren Module verwenden ASCII, damit der VBA-Import unabhängig von der Windows-Codepage funktioniert; die Oberfläche der Vorlage verwendet Unicode.

Die Versionsgrenzen und die Ergebnisse der Cloud-Prüfungen stehen in [docs/Validierung.md](docs/Validierung.md). Pumpenkennlinien, Regelventile, mehrere Rohrleitungsabschnitte und Zweiphasenmodelle sind Erweiterungen und noch nicht implementiert.
