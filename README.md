# Druckverlust – Excel/VBA-Systemkennlinie

**V2-Entwicklung auf `develop/network-v2`:** Die Bestandsaufnahme sowie Netzwerkdatenmodell und Solver (Phasen 1–3) sind separat ergänzt. Einstieg: [Bestandsaufnahme](docs/v2/Bestandsaufnahme.md), [Installation und Phasenstand](docs/v2/Netzwerk.md), [Validierung](docs/v2/Validierung.md). Neue Vorlage: `excel/Druckverlust_Netzwerk_V2.xlsx`; neue Module: `vba/network/`. Die nachfolgend beschriebene bisherige Version einschließlich aller acht Module, Standalone-Dateien und Tests bleibt unverändert. Schemagrafik, Netzwerk-Systemkennlinie und Netzwerkbericht sind noch nicht Teil dieses Entwicklungsstands.


Version 1.2 für inkompressible, einphasige Newtonsche Flüssigkeiten mit konstanten Stoffwerten und einem einheitlichen Hauptrohr-Innendurchmesser, ergänzt um eine Lochblende und eine Pumpenkennlinie. Berechnet Darcy-Weisbach-Rohrreibung, Einzelwiderstände, geodätische Höhe und einen konstanten Behälterdifferenzdruck.

## Installation in Microsoft Excel für Windows

1. `excel/Druckverlust_Vorlage.xlsx` lokal öffnen und **Speichern unter → Excel-Arbeitsmappe mit Makros (*.xlsm)** wählen, beispielsweise `Druckverlust.xlsm`.
2. Mit **Alt+F11** den VBA-Editor öffnen. Im Projekt-Explorer das Projekt **dieser Arbeitsmappe** auswählen.
3. Über **Datei → Datei importieren** alle acht Dateien aus `vba/` importieren: `Hydraulik.bas`, `Eingaben.bas`, `Oberflaeche.bas`, `Validierung.bas`, `Pumpenmodell.bas`, `Betriebsoberflaeche.bas`, `ErweiterungsTests.bas`, `EngineeringBericht.bas`. Keine Module in PERSONAL.XLSB importieren. Vorhandene Module derselben Namen vor einer Aktualisierung gezielt ersetzen.
4. **Debuggen → VBAProject kompilieren** ausführen. Es sind keine zusätzlichen Verweise oder Add-ins erforderlich.
5. Zurück zu Excel; **Alt+F8 → Einrichtung → Ausführen**. Dies legt sieben Schaltflächen auf `Eingabe` an. Wiederholtes Ausführen ersetzt nur die eigenen Schaltflächen.
6. Speichern. Makros nur für die geprüfte Arbeitsmappe gemäß der eigenen Unternehmensrichtlinie freigeben. Eine generelle Absenkung der Excel-Makrosicherheit ist nicht erforderlich.
7. **Validierung starten** ausführen. Erwartet sind zwei Erfolgsmeldungen: **37 Erweiterungs-Testfälle** und **20 Rechenkern-Testfälle bestanden**. Anschließend **Systemkennlinie berechnen** anklicken.

Eine fertige XLSM wird nicht mitgeliefert: In der Cloud steht kein Microsoft Excel zur Verfügung, um ein VBA-Projekt einzubetten, zu kompilieren und auszuführen. Die XLSX enthält absichtlich keine scheinbar bereits berechneten Resultate. Nach Import und Speichern läuft die Anwendung vollständig lokal ohne Python oder externe Dienste.

## Bedienung

- Blaue Zellen auf `Eingabe` bearbeiten. DN wählen und bei Bedarf **DN-Vorschlag übernehmen** drücken. Dies ersetzt B6 ausdrücklich; anschließend den tatsächlichen Innendurchmesser anhand Rohrnorm und Wandstärke überprüfen. Die Berechnung verwendet ausschließlich B6, nicht DN.
- Länge und Rauigkeit dürfen null sein; Durchmesser, Dichte und Viskosität müssen positiv sein. 0 ≤ Qmin < Qmax; 2–10000 ganzzahlige Berechnungspunkte. Positive Höhe bedeutet höher gelegenen Austritt. Positiver Behälterdruck bedeutet höherer Druck im Zielbehälter.
- In der Armaturentabelle den Typ wählen und eine ganzzahlige Anzahl eingeben. **Zeta manuell leer** verwendet den aktuellen Datenbankwert, eine Zahl einschließlich 0 überschreibt ihn. Nach Typwechsel bleibt eine manuelle Überschreibung bewusst erhalten; zum Verwenden des neuen Richtwerts die Zelle leeren. Aktive freie Widerstände benötigen eine explizite Eingabe.
- Weitere Zeilen mit **Armatur hinzufügen** anlegen. Anzahl 0 deaktiviert eine Komponente. Die Tabelle kann über die Excel-Tabellenfunktionen verkleinert werden; ihre Namen und die ersten drei Spalten nicht ändern.
- **Nach jeder Eingabe-, Tabellen- oder Datenbankänderung neu berechnen.** Es gibt keine automatisch laufenden Änderungsereignisse. Die Berechnung aktualisiert Ergebnistabelle und alle drei XY-Kurven in einem Lauf. `Ergebnisse!A2` zeigt den Laufzeitpunkt und die verwendete Summe n·ζ. Die Diagrammkurven sind für gerenderte Berechnungspunkte verbunden, keine Interpolation des Strömungsmodells.
- Warnungen stehen pro Punkt in Spalte K; der Übergangsbereich wird zusätzlich farblich markiert. Bei ungültigen Eingaben werden vorhandene Ergebnisse als **nicht aktuell** markiert, nicht als neuer Lauf ausgegeben.
- Datenbank-Richtwerte sind ausdrücklich Näherungen. Für eine technische Auslegung geeignete Herstellerwerte und Randbedingungen einsetzen. Werte beziehen sich auf die Geschwindigkeit im eingegebenen Rohrdurchmesser.

## Neu in Version 1.2

**Vorhandene V1-Anwender:** Die neue XLSX-Vorlage verwenden, als neue XLSM speichern und **alle acht Module** importieren. Alte Eingabewerte gezielt übertragen. Ein reiner Modulimport in die alte Arbeitsmappe reicht nicht, weil neue Blätter benötigt werden. Keine angepasste alte Arbeitsmappe überschreiben.

### Lochblende

Auf `Lochblende` Aktiv=1 setzen, DN wählen, über **Blenden-DN übernehmen** auf `Eingabe` den Di-Vorschlag einsetzen und tatsächlichen Di sowie Lochdurchmesser prüfen. Cd (Startwert 0,61) ist überschreibbar. Zulässiges Durchmesserverhältnis d/Di: 0,2 bis 0,75. Die Anzahl gilt für gleiche, hydraulisch voneinander unabhängige Blenden; nahe hintereinander montierte Blenden sind damit nicht abgebildet.

Berechnet wird der **bleibende Druckverlust nach Druckrückgewinnung**, nicht der Messdifferenzdruck über den Druckabgriffen. Er ist in den Einzelwiderständen (Spalte F) enthalten; Spalte L weist ihn zusätzlich als Teilbetrag aus. Nicht nochmals addieren. Abweichender lokaler Rohr-Di ist möglich; Übergangs-/Reduzierverluste müssen separat erfasst werden. Siehe [Modell und Grenzen](docs/Erweiterungen.md).

### Einzelner Betriebspunkt bei vorgegebenem Q

Q in `Eingabe!G3` in m³/h eingeben und **Einzelpunkt berechnen** wählen. Alle aktuellen Rohr-, Stoff-, Armaturen-, Kv- und Blendenwerte werden gelesen; der Kennlinienbereich Qmin/Qmax ist dafür nicht maßgeblich. Ergebnis: `Betriebspunkte`, Zeile 5. Dies bestimmt den erforderlichen Systemdruck bei Q; es ist keine Pumpenschnittpunktsuche.

### Pumpenkennlinie und Betriebspunkt

Auf `Pumpe` die **illustrativen Beispielwerte** in `tblPumpe` durch Herstellerdaten ersetzen: Q in m³/h und Förderhöhe H in m der geförderten Flüssigkeit. Tabellenzeilen mit Tab am Tabellenende hinzufügen; nicht benötigte Zeilen über **Tabellenzeilen löschen** entfernen. Mindestens 2, höchstens 1000 Punkte, keine Leerzeilen, Q streng steigend, H streng fallend. Die Druckspalte C wird berechnet und ist keine Eingabe. Kennlinien mit Buckel, Plateaus oder mehreren möglichen Schnittpunkten werden in dieser Version ausdrücklich abgelehnt.

**Pumpenbetriebspunkt** auf `Eingabe` starten. Der Solver verwendet die gesamte eingegebene Pumpentabelle, unabhängig von Qmin/Qmax der Systemkurve, ohne Extrapolation. Ergebnis: `Betriebspunkte`, Zeile 11, einschließlich erforderlichem Systemdruck, Verlustanteilen, Warnungen und Druckresiduum (N11/O11). Auf `Pumpe` erscheint ein Diagramm mit Pumpenkennlinie, Systemkurve und markiertem Schnittpunkt. Kein Schnittpunkt im Tabellenbereich führt zu einer Meldung.

Die drei Modi (Systemkurve, Einzelpunkt, Pumpenschnittpunkt) besitzen eigene Laufzeitangaben. Jeder Modus muss nach Änderungen erneut gestartet werden; andere Ergebnisse werden dadurch nicht automatisch neu berechnet.

## Kv-Armatur und Engineering-Dokumentation (Version 1.2)

**Kv-Armatur:** Auf `KvArmatur` Aktiv=1 setzen, den aktuellen Kv in m³/h und die Anzahl gleicher Armaturen **in Serie** eingeben. Der Kv muss zum Öffnungsgrad passen; Kvs ist nur bei voller Öffnung geeignet. Der Verlust wird in Systemkurve, Einzelpunkt und Pumpenschnittpunkt berücksichtigt. Spalte F enthält sämtliche Einzelverluste, Spalte M zusätzlich den bereits enthaltenen Kv-Anteil. Nicht doppelt addieren.

**Eigenständige Kv-Berechnung:** `excel/Kv_Standalone.xlsx` enthält das makrofreie Blatt `Kv_Standalone`; dasselbe Blatt ist zusätzlich in `Druckverlust_Vorlage.xlsx` enthalten. Q, Kv bei aktueller Öffnung, Dichte und Anzahl gleicher Armaturen in Serie eingeben. Viskosität und optionale Temperatur werden dokumentiert, ohne automatische Korrektur. Ausgegeben werden Einzel- und Gesamtdruckverlust in bar, Gesamtdruckverlust in Pa sowie Verlusthöhe in m. Projekt, Dokument, Bearbeiter und Revision ergänzen; Druckbereich ist **A4 Hochformat, eine Seite** mit sichtbaren Formeln, Annahmen und Freigabefeld. Die Berechnung ist unabhängig von `KvArmatur` und den Systemeingaben; Werte für eine Systemberechnung dort separat eintragen. Bei ungültigen Eingaben werden Ergebnisse als #N/A markiert.

**Eigenständige Blendenberechnung:** Die separate Datei **`excel/Drosselblende_Standalone.xlsx`** enthält nur das Blendenblatt. Auch das Blatt `Blende_Standalone` in der Gesamtvorlage benötigt weder Makros noch Werte anderer Blätter. DN, tatsächlichen Di, Lochdurchmesser, Q, Dichte, Viskosität, Cd und optional Temperatur direkt eingeben. Der DN-Di-Vorschlag rechts ist eine Hilfe; der tatsächliche Di wird manuell eingetragen. Ergebnisse und Gültigkeitsstatus aktualisieren sich über Excel-Formeln. Das Blatt lässt sich allein in eine andere Arbeitsmappe kopieren. Projekt, Dokumentnummer, Bearbeiter und Revision ausfüllen; anschließend direkt drucken oder als PDF exportieren: **A4 Hochformat, eine Seite**, einschließlich Formeln, Modellgrenzen und Freigabefeld.

**Systembericht:** Auf `Systembericht` Projekt-/Dokumentdaten ausfüllen und `D5` auf 0 (ohne Pumpe) oder 1 (mit eingetragener Pumpenkennlinie) setzen. Nach `Einrichtung` befindet sich rechts neben dem Druckbereich die Schaltfläche **Systembericht erstellen**; alternativ Alt+F8 → `SystemberichtErstellen`. Der Bericht berechnet Systemkurve und optional Pumpenschnittpunkt **neu aus den aktuellen Eingaben**. Bei Pumpeneinschluss wird der Diagrammbereich gegebenenfalls bis zu den Pumpentabellen-Endpunkten erweitert. Ungültige Pumpendaten oder fehlender Schnittpunkt ergeben eine Fehlermeldung; alte Daten werden als nicht aktuell markiert.

Der Ausdruck enthält ein A4-Deckblatt mit Diagramm, Formeln, Annahmen und optional Betriebspunkt sowie einen mehrseitigen Anhang mit Eingabestand, wirksamen Armaturenwerten, Kv-/Blendenparametern, verwendeten Pumpenstützpunkten und Warnungen. Die Diagrammdaten werden im Berichtsblatt selbst gespeichert. Nach Änderungen **Bericht neu erstellen**; der Bericht ist ein dokumentierter Stand, keine laufend aktualisierte Verknüpfung. Druckbereich und Seitenumbruch sind voreingestellt. Die Excel-Druckvorschau und der VBA-Lauf sind vor Engineering-Freigabe auf Windows zu prüfen.

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
tools/build_extension_tests.py   Erzeugung der Blenden-/Pumpenprüfungen
vba/Pumpenmodell.bas             Blendenintegration, Pumpeninterpolation, Nullstellensuche
vba/Betriebsoberflaeche.bas       Einzelpunkt, Pumpenbetrieb und Pumpendiagramm
vba/ErweiterungsTests.bas        Zusätzliche VBA-Selbsttests
vba/EngineeringBericht.bas      Neuberechnung, Eingabesnapshot und Bericht
tools/report_sheets.py          Drucklayouts und eigenständige Excel-Formeln
```

Entwicklung benötigt Python 3.10+ und, ausschließlich zum Erzeugen der Vorlage, `XlsxWriter==3.2.9`:

```bash
python -m pip install -r requirements-build.txt
python tools/build_workbook.py
python tools/build_validation.py
python tools/build_extension_tests.py
python -m unittest discover -s tests -v
```

Diese Befehle aus dem Repository-Verzeichnis ausführen. Die drei Generatoren überschreiben ausschließlich ihre eigenen generierten Artefakte: XLSX, Referenz-JSON und die beiden Validierungs-Module. Änderungen an einer lokal angepassten Arbeitsmappe vorher separat sichern. Die importierbaren Module verwenden ASCII, damit der VBA-Import unabhängig von der Windows-Codepage funktioniert; die Oberfläche der Vorlage verwendet Unicode.

Die Versionsgrenzen und die Ergebnisse der Cloud-Prüfungen stehen in [docs/Validierung.md](docs/Validierung.md). Lochblende, Kv-Armatur, Einzelpunkt, Pumpenschnittpunkt und druckfertige Engineering-Blätter sind implementiert. Regelventile, mehrere Rohrleitungsabschnitte und Zweiphasenmodelle sind noch nicht implementiert.
