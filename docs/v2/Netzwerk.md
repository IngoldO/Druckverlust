# V2 – Datenmodell und Netzwerklöser (Phasen 2 und 3)

## Status

| Phase | Stand |
|---|---|
| 1 Bestandsaufnahme | Dokumentiert, Rückkehrstand und SHA-256-Inventur vorhanden |
| 2 Datenmodell | Knoten, aktive Leitungen, individuelle Armaturen, getrennte Stoffwerte und Netzpumpe implementiert |
| 3 Solver | Gedämpftes gemischtes Newton-Verfahren, separate Residuen, Null-/Rückströmung, Maschen, Rückschlaglogik, Fehlermeldungen implementiert; Cloud-Verfahrensprüfung bestanden |
| 4 Excel-Oberfläche | Grundlegende Tabellen/Dropdowns und Berechnungsmakro vorhanden; weiterführende Bedienung und Zusammenfassungsübersicht offen |
| 5 Netzwerkschema | Offen; X/Y werden bereits gespeichert, noch keine Shape-Darstellung |
| 6 Pumpe/Systemkurve | Eine Pumpe bereits direkt in Netzwerkgleichungen gekoppelt; Netzwerk-Systemkennlinien-Sweep mit auswählbarer Einspeisung noch offen |
| 7 Netzwerkbericht | Offen; bisheriger Einzelrohr-Systembericht und Standalone-Berichte unverändert |
| 8 Abnahme | 57 Cloud-Tests bestanden; tatsächliche VBA-Ausführung, Migration einer externen V4-XLSM und Windows-Excel-Abnahme offen |

Die bisherige Einzelrohr-Systemkurve bleibt vollständig vorhanden. Sie ist keine Ersatzberechnung für die noch offene allgemeine Netzwerk-Systemkurve.

## Installation und bewusste Übernahme vorhandener Werte

1. Alte Arbeitsmappe unverändert aufbewahren. `excel/Druckverlust_Netzwerk_V2.xlsx` als **neue** XLSM speichern.
2. Alle acht bisherigen BAS-Dateien aus `vba/` und zusätzlich alle **fünf** BAS-Dateien aus `vba/network/` importieren. Keine neuen Module ersetzen alte Namen.
3. VBA-Projekt kompilieren. Bisheriges `Einrichtung` für Modus A ausführen. Neue Makros sind unabhängig über Alt+F8 zugänglich: `NetzwerkValidieren`, `NetzwerkBerechnen`.
4. `ValidierungStarten` muss weiterhin 37 Erweiterungs- und 20 Kernfälle melden. `NetzwerkValidieren` erwartet 25 zusätzliche Netzwerkfälle. Alle diese Makros sind vorbereitet, nicht hier in Excel ausgeführt.
5. Fiktive Werte in `Netzwerk_*` bearbeiten und `NetzwerkBerechnen` starten. Der Ergebnisstatus muss KONVERGIERT lauten; separate Residuen prüfen.
6. Alte Eingabewerte bei Bedarf bewusst in die entsprechenden bisherigen Blätter der neuen Datei übertragen. Keine automatische Migration, kein Überschreiben von Anwenderdateien. Für eine externe V4-XLSM muss zuerst deren konkreter Blatt-/Modulbestand geprüft werden.

Alternative für eine bestehende XLSM: eine Kopie verwenden; nur die sechs neuen Netzwerkblätter aus der V2-Vorlage kopieren, anschließend ausschließlich die fünf neuen Module importieren. Die bestehenden Tabellen `tblDatenbank` und der Name `DNWerte` müssen vorhanden sein. Die strukturierten Namen `NetzKnotenIDs` und `NetzLeitungsIDs` auf die kopierten Tabellen prüfen. Diese Variante ist ebenfalls erst nach Kompilieren und Regression in Windows-Excel freizugeben.

Die makrofreien Standalone-Dateien werden weder verändert noch für Modus B benötigt. `Blende_Standalone` und `Kv_Standalone` bleiben auch in der V2-Gesamtvorlage enthalten.

## Tabellen und Einheiten

### Netzwerk_Setup

B5: Dichte kg/m³; B6: dynamische Viskosität mPa·s; B7: Iterationslimit (1–1000); B8: maximale freie Knotenbilanz m³/h; B9: maximale Energieverletzung bar. Diese Werte sind von `Eingabe` unabhängig. Standard: 1000 kg/m³, 100 mPa·s, 100 Iterationen, 10⁻⁶ m³/h und 10⁻⁸ bar. Die hohe Viskosität gehört zum fiktiven laminaren Startbeispiel, nicht zu einer allgemeinen Stoffempfehlung.

### Netzwerk_Knoten / tblNetzKnoten

ID, Bezeichnung, Typ, Höhe m, optionaler Druck bar(a), optionales externes Q m³/h, optionale X/Y-Positionen. IDs sind nicht leer, Groß-/Kleinschreibung unterscheidet keine IDs. Die Typen Einspeisung, Verteiler, Verbindungsknoten, Sammelkollektor und Verbraucher beschreiben die Rolle; Randbedingungen werden explizit gesetzt. Behälter und offener Auslass benötigen einen positiven Absolutdruck. Lokaler Atmosphärendruck wird nicht stillschweigend angenommen.

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

Vorhandene Einzelrohr-, Pumpen- und Engineering-Funktionen sind unverändert. Noch keine automatischen Netzwerk-Shapes, kein Netzwerkbericht, keine allgemeine Netzwerk-Systemkennlinie und keine vollständige Netzwerk-Kennzahlenübersicht. X/Y sowie Knoten-/Leitungsbezeichnungen sind bereits für die folgenden Phasen vorgesehen. Stoff-/Modellgrenzen aus V1 gelten weiter; keine Kavitation-/Phasen-/NPSH-/Thermo-/Transientenprüfung. Konstante ζ-/Kv-/Cd-Werte bei kleinem Re sind besonders unsicher; die laminaren fiktiven Validierungsfälle prüfen die Gleichungen, nicht die Auslegungsgültigkeit solcher Kennwerte.
