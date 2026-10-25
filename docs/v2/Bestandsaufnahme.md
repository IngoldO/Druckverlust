# V2 – Phase 1: Bestandsaufnahme und Schutz der bisherigen Version

Ausgangspunkt: Git-Commit `e8bd020cdb471f23b4a68f344a5861e660ba5725` auf `main`. Entwicklung ausschließlich auf `develop/network-v2`. Rückkehr: bisherigen Stand auf `main` verwenden oder diesen Commit separat auschecken; keine Rücksetzung einer Arbeitsmappe mit eigenen Eingaben durchführen. Benutzerdateien vorher unabhängig sichern.

## Tatsächlich vorhandene Dateien

Im ausgewählten Repository liegt **keine XLSM-Datei**, insbesondere keine `Druckverlust_Vorlage_V4.xlsm`. Vorhanden sind `Druckverlust_Vorlage.xlsx`, `Drosselblende_Standalone.xlsx` und `Kv_Standalone.xlsx`. Keine enthält `xl/vbaProject.bin`. Die vorhandene README beschreibt diese Dateien korrekt: Makros werden durch Import der BAS-Dateien ergänzt und anschließend als XLSM gespeichert. Eine möglicherweise außerhalb dieses Repository-Stands bereitgestellte V4-XLSM konnte nicht untersucht werden; ihr VBA-Projekt wird nicht erfunden oder überschrieben.

Die maschinenlesbare vollständige Bestandsaufnahme steht in `Bestand_V1.json`: SHA-256 der bisherigen Module, Generatoren, Tests und Arbeitsmappen; öffentliche Makros/Funktionen; Blätter, Tabellen, Diagramme, Namen und Testnamen. Sie wurde vor der Implementierung erstellt und wird nicht nachträglich aktualisiert.

## Funktionsbestand

- Hydraulik: Darcy-Weisbach, Colebrook, laminar/Übergang/turbulent, Re, Rauigkeit, statische Höhe, Behälterdruck, konstante Stoffwerte, Q=0 und Systemkennlinie.
- Armaturen: dynamische Typ-/Anzahltabelle mit Datenbank, explizitem ζ=0, freien Widerständen, individuellen Kv-/Blendenmodellen und separat ausgewiesenen Teilverlusten ohne Doppelzählung.
- Pumpe: fallende Q/H-Herstellerkennlinie, stückweise lineare Interpolation, Dichtebezug, keine Extrapolation, Schnittpunkt mit Residuenkontrolle. Einzelpunkt bei vorgegebenem Q.
- Dokumentation: Systembericht mit Projektdaten, Dokumentnummer, Revision, Freigabe, Eingabesnapshot und Diagrammen. Makrofreie, lokale Kv-/Blendenformeln mit A4-Drucklayout.

### Zehn bisherige Blätter

`Eingabe`, `Ergebnisse`, `Armaturendatenbank`, `Lochblende`, `Pumpe`, `Betriebspunkte`, `KvArmatur`, `Blende_Standalone`, `Systembericht`, `Kv_Standalone`.

Tabellen: `tblArmaturen`, `tblDatenbank`, `tblPumpe`. Gemeinsame Namen: `DNWerte`, `ArmaturenTypen`; zusätzlich lokale Druckbereiche. Zwei vorbereitete XY-Diagramme in der XLSX; weitere Diagramme/Shapes werden von VBA zur Laufzeit erzeugt. Die JSON-Inventur unterscheidet Paketinhalt von nur im Code vorhandenen Laufzeitfunktionen.

### Acht bisherige VBA-Module

`Hydraulik`, `Eingaben`, `Oberflaeche`, `Pumpenmodell`, `Betriebsoberflaeche`, `EngineeringBericht`, `Validierung`, `ErweiterungsTests`.

Bedienmakros: `Einrichtung`, `SystemkennlinieBerechnen`, `DNAnwenden`, `ArmaturHinzufuegen`, `BlendenDNAnwenden`, `EinzelpunktBerechnen`, `PumpenbetriebspunktBerechnen`, `SystemberichtErstellen`, `ValidierungStarten`, `ErweiterungenValidieren`. Zusätzliche öffentliche Helfer sind vollständig im JSON aufgeführt.

### Bestehende Prüfungen

38 Python-Tests bestanden vor Änderungen. 20 Rechenkern- und 37 Erweiterungsfälle sind als tatsächliche VBA-Testmakros vorhanden, wurden mangels Microsoft Excel aber nicht in der Cloud ausgeführt. Diese Unterscheidung bleibt bestehen.

## Erweiterungsstrategie

Die bisherigen BAS-Dateien, Tests, Generatoren und drei Arbeitsmappen bleiben **bytegleich**. Regressionstests prüfen ihre gespeicherten SHA-256-Werte. V2-Module liegen separat unter `vba/network`, weitere Testwerkzeuge unter `tools/network`. Modus A bleibt unverändert; Modus B verwendet eigene Tabellen und eigene Stoffwerte. Derselbe unveränderte Hydraulikkern berechnet lokale Verlustfunktionen.

`Druckverlust_Netzwerk_V2.xlsx` wird mit einem separaten Generator erzeugt. Dieser ruft den ursprünglichen Generator nur mit einem temporären Ausgabeverzeichnis auf und ergänzt neue Blätter vor dem Schließen. Alte Dateien werden nicht neu gespeichert. Tests vergleichen Zellwerte/Formeln/Formatindizes sowie die bisherigen Diagramm- und Tabellen-XMLs der neuen Vorlage mit V1.

Phasenstand und Anleitungen: [Netzwerk.md](Netzwerk.md). Dies ist die zunächst beauftragte Implementierung der Phasen 1–3 mit einem grundlegenden Excel-Ein-/Ausgabeadapter. Ein vollständig abgenommener V2-Netzwerksimulator mit Schemagrafik und Netzwerkbericht wird damit noch nicht behauptet.
