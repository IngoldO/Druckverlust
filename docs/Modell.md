# Modell, Systemgrenzen und Datenqualität

## Gleichungen und Einheiten

Die Eingabe erfolgt in m³/h, mm, m, kg/m³, mPa·s und bar. Der Rechenkern übernimmt Q in m³/h, D, L, ε und Δh in m, μ in Pa·s und den konstanten Behälterdifferenzdruck in bar. Er konvertiert Q/3600 zu m³/s und Druck in Pa durch Division durch 100000 zu bar. g = 9,80665 m/s².

- A = πD²/4; v = Q/A (mit Q in m³/s).
- Re = ρvD/μ.
- Δp_R = λ(L/D)ρv²/2 (Darcy-, nicht Fanning-Reibungsbeiwert).
- Δp_E = Σ(nᵢζᵢ)ρv²/2.
- Δp_H = ρg(z_Ziel − z_Quelle).
- Δp_System = Δp_R + Δp_E + Δp_H + (p_Ziel − p_Quelle).

Die Systemkennlinie beschreibt die erforderliche Druckerhöhung zwischen den angegebenen Grenzen, ohne Pumpe. Vorzeichen gelten für die positive Richtung Quelle → Ziel; negative Volumenströme werden abgelehnt. Negative Gesamtdrücke bleiben erhalten und werden als statischer Antrieb gekennzeichnet. Für Reservoirgrenzen müssen Eintritt/Austritt berücksichtigt werden. Bei gleicher Rohrweite an beiden Rohrgrenzen heben sich die dortigen kinetischen Energieanteile auf; zusätzliche Reservoirverluste nicht doppelt ansetzen.

## Reibungsbeiwert

Für 0 < Re < 2300 gilt λ = 64/Re. Für Re ≥ 4000 wird die implizite Colebrook-White-Gleichung gelöst:

```text
1/√λ = −2 log10(ε/(3,7 D) + 2,51/(Re √λ))
```

Der Solver iteriert x = 1/√λ ab x₀ = 7, höchstens 200 Schritte. Abbruch bei |x_neu − x_alt| < 10⁻¹⁰. Ein fehlender Abbruch erzeugt einen Fehler, keinen unkontrollierten Näherungswert. Dies ist eine numerische Lösung der Gleichung, keine zusätzliche explizite Korrelation. Die in der Cloud untersuchten Re/ε-D-Kombinationen weichen beim Reibungsbeiwert weniger als 10⁻¹¹ absolut von unabhängiger Bisektion ab. Diese numerische Übereinstimmung ist **keine physikalische Genauigkeitszusage**.

Für 2300 ≤ Re < 4000 wird max(64/Re, λ_Colebrook) verwendet. Colebrook wird hierfür außerhalb seines gesicherten turbulenten Bereichs ausgewertet. Das Ergebnis ist nur eine konservativ gewählte Rechenkonvention zwischen diesen beiden Ansätzen, **keine garantierte obere Schranke realer Verluste**. Jede betroffene Zeile wird als UNSICHER markiert. Eine unstetige Modelländerung bei Re=2300 ist bewusst möglich; keine unkommentierte Interpolation. Auslegung in diesem Bereich erfordert eigene Untersuchung.

Bei Q=0 sind v=Re=Δp_R=Δp_E=0; λ=0 ist ausschließlich eine Rechenkonvention, kein physikalischer Reibungsbeiwert. Statische Beiträge bleiben erhalten.

## Gültigkeit

Für die Hauptrohrstrecke (Erweiterungen siehe [Lochblende und Pumpe](Erweiterungen.md)): einphasige, inkompressible Newtonsche Flüssigkeit; konstante Dichte/Viskosität; stationäre, hinreichend ausgebildete Strömung; runder, vollständig gefüllter Querschnitt; überall derselbe Innendurchmesser. Relative Rauigkeit 0 ≤ ε/D ≤ 0,05; höhere Werte werden abgelehnt. Re≥4000 ist die übliche Anwendungsgrenze des hier verwendeten turbulenten Ansatzes. Sehr hohe Reynolds-Zahlen und die Übertragbarkeit äquivalenter Sandrauigkeit sind projektbezogen zu beurteilen.

Temperatur ist nur dokumentierter Betriebsparameter; keine automatische Stoffwertkorrektur. Außer Werten unter dem absoluten Nullpunkt kann die Software aus Temperatur allein weder die Phase noch Stoffeigenschaften prüfen. Kavitation, Dampfdruck, Sieden, Zweiphasigkeit, Kompressibilität, nicht-Newtonsche Effekte, Wärmeübertragung, Druckstöße, Einlaufstrecken und Rohrverformung werden nicht modelliert. Absolutdruck und Dampfdruck sind keine Eingaben; eine Kavitationserkennung wäre daher nicht sachgerecht. v > 5 m/s erzeugt einen Hinweis, keine universelle Geschwindigkeitsgrenze.

## Armaturendaten

Die konkreten Werte sind **illustrative Startwerte**, keine normativen oder editionsgenau zitierten Herstellerdaten. Grundlage ist die übliche Größenordnung der Einzelwiderstandsmethode; die Tabelle stellt keine quantitative Auswertung der folgenden Fachbücher dar. Für Nachrechnung und Auswahl geeigneter Kennwerte:

- Crane Co., *Flow of Fluids Through Valves, Fittings, and Pipe*, Technical Paper No. 410 (TP-410).
- I. E. Idelchik, *Handbook of Hydraulic Resistance*, 3. Auflage, CRC Press, 1994.
- C. F. Colebrook, *Turbulent Flow in Pipes, with Particular Reference to the Transition Region Between the Smooth and Rough Pipe Laws*, Journal of the Institution of Civil Engineers 11(4), 1939, S. 133–156; DOI: 10.1680/ijoti.1939.13150.

| Komponente | ζ Startwert | Einschränkung / Datenqualität |
|---|---:|---|
| 90°-Bogen | 0,9 | Beispiel; Radius, Segmentierung, Rauigkeit prüfen |
| 45°-Bogen | 0,4 | Beispiel; Radius und Bauart prüfen |
| T-Stück gerader Durchgang | 0,6 | Beispiel; Durchflussaufteilung prüfen |
| T-Stück Abzweig | 1,8 | Beispiel; 90°, Aufteilung und Flächen prüfen |
| Kugelhahn | 0,05 | Beispiel; voll geöffnet, voller Durchgang |
| Absperrschieber | 0,15 | Beispiel; voll geöffnet |
| Sitzventil | 10 | Grobe Annahme, voll geöffnet; starke Bauformabhängigkeit |
| Rückschlagventil | 2 | Grobe Annahme für Klappenbauart; Öffnung und Einbau prüfen |
| Schmutzfänger | 2 | Grobe Annahme für sauberen Zustand; Sieb und Verschmutzung prüfen |
| Eintritt | 0,5 | Idealisierter scharfkantiger Eintritt aus großem Reservoir |
| Austritt | 1 | Dissipation der Geschwindigkeitsenergie im großen Reservoir |
| Frei | 0 | Platzhalter; für aktive Zeile manuell festlegen |

Insbesondere T-Stücke benötigen üblicherweise Durchflussverhältnisse, die Version 1 nicht erfasst. Bei kleinem Re sind konstante ζ-Werte häufig ungeeignet; bei Re<4000 und aktiven Einzelwiderständen erscheint ein zusätzlicher Hinweis. Alle ζ beziehen sich auf dieselbe Rohrgeschwindigkeit; Herstellerwerte mit anderer Bezugsfläche müssen umgerechnet werden.

Die DN-Tabelle ist eine illustrative Stahlrohr-Abmessungsreihe ohne verbindliche Norm-/Schedule-Zuordnung. DN legt keinen eindeutigen Innendurchmesser fest. Der explizite Übernahmeknopf verhindert, dass eine bewusst manuell eingegebene lichte Weite bei der Berechnung überschrieben wird.

## Erweiterungspunkte

`Hydraulik` kennt weder Arbeitsblätter noch Diagramme und gibt einen strukturierten Betriebspunkt zurück. `Eingaben` liest Tabellen und überprüft Eingaben; `Oberflaeche` orchestriert und zeichnet. Spätere Abschnitte können eigene Durchmesser und Stoffwerte erhalten und ihre Verlustanteile addieren. `Pumpenmodell` ergänzt die Kennlinieninterpolation, Blendenintegration und Schnittpunktsuche; `Betriebsoberflaeche` übernimmt deren Ausgabe. Die aktuellen Ergänzungen sind in [Erweiterungen.md](Erweiterungen.md) dokumentiert. Regelventilmodelle benötigen weitere Datenstrukturen; Zweiphasenströmung benötigt ein anderes physikalisches Modell, nicht lediglich einen Korrekturfaktor.
