# Lochblende, Einzelpunkt und Pumpenschnittpunkt

## Lochblende

Geometrie: dünne, scharfkantige, konzentrische Blende in einem runden, vollständig gefüllten Rohr; inkompressible Flüssigkeit. D_B ist der lokale Rohrinnendurchmesser, d der Lochdurchmesser, β=d/D_B. DN dient nur als Auswahlhilfe; gerechnet wird mit dem überprüften tatsächlichen D_B.

Für einen konstanten, vom Benutzer vorzugebenden Durchflussbeiwert C_d wird der bleibende Verlustbeiwert angesetzt:

```text
A = sqrt(1 − β⁴ (1 − C_d²))
ζ_B = (A / (C_d β²) − 1)²
Δp_B = n ζ_B ρ v_B² / 2
v_B = (Q/3600) / (π D_B²/4)
```

Dies entspricht der üblichen Druckrückgewinnungsbeziehung für Blenden, wie im Zusammenhang mit ISO 5167-2 behandelt. Anders ausgedrückt:

```text
Δp_Mess = (1 − β⁴)/(C_d² β⁴) · ρv_B²/2
Δp_bleibend / Δp_Mess = (A − C_d β²)/(A + C_d β²)
```

**Keine Umsetzung einer normgerechten ISO-5167-Messstelle:** Es gibt keine automatische Reader-Harris/Gallagher-Cd-Korrelation, keine Druckabgriff-, Dicken- oder Einlaufstreckenberechnung. C_d=0,61 ist ein illustrativer Ausgangswert, nicht universell gültig. Für technische Auslegung Herstellerdaten oder eine geeignete geometrie- und Reynolds-abhängige Korrelation einsetzen. Der Bezug auf ISO 5167-2 beschreibt den Zusammenhang der Druckrückgewinnung; er ist keine Konformitätsaussage.

Implementierter Geometriebereich: 0,2≤β≤0,75 und 0<C_d<1. Bei β-Grenzen wird ein rein numerischer Rundungsspielraum von 10⁻¹² verwendet. Bei lokalem Re_B<10000 wird ausdrücklich gewarnt. Diese Warnschwelle ist keine allgemeine Gültigkeitszusage oberhalb von 10000. Mehrere Blenden dürfen nur als unabhängige Verluste addiert werden, wenn hinreichende Druckrückgewinnung zwischen ihnen gegeben ist. Kompakte mehrstufige Drosselpakete sind nicht modelliert.

Zur Integration in die bisherige Systemgleichung wird auf die Hauptrohrgeschwindigkeit umgerechnet:

```text
ζ_äquivalent = n ζ_B (D_Haupt / D_B)⁴
```

Dadurch bleibt der physikalische Blendenverlust bei abweichendem lokalem D_B korrekt auf dessen Geschwindigkeit bezogen. Die übrigen Rohrstrecken behalten D_Haupt; zusätzliche Rohrstücke, Reduzierungen oder Aufweitungen müssen separat berücksichtigt werden. Der Blendenanteil erscheint als Teil der Einzelwiderstände und zusätzlich separat in Spalte L. Die Summe darf ihn nur einmal enthalten.

Q=0 ergibt auch mit Blende null dynamischen Verlust. Absolutdruck/Dampfdruck werden weiterhin nicht erfasst; Kavitation, Flashing, Geräusch, Erosion und zulässige Belastung der Blende sind damit nicht geprüft.

## Einzelpunkt

`EinzelpunktBerechnen` liest dieselben Daten wie die Systemkurve über `LeseAnlage` und ruft denselben `Anlagenpunkt` auf. Nur der gewünschte Volumenstrom kommt aus `Eingabe!G3`. Eine Pumpenkennlinie und ein gültiger Kennlinien-Abtastbereich sind für diesen Modus nicht nötig. Alle Verlustanteile, statischen Beiträge, Regime und Hinweise werden ausgewiesen.

## Pumpe

Eingaben sind zusammengehörige Paare (Q_i,H_i) bei einer festen Drehzahl und einem festen Laufraddurchmesser. Voraussetzung: Q_i≥0 streng steigend, H_i≥0 streng fallend. Es werden 2–1000 Punkte unterstützt. H ist die Förderhöhe der geförderten Flüssigkeit; Umrechnung in Druck:

```text
Δp_Pumpe(Q) = ρ g H(Q) / 100000    [bar]
```

H(Q) wird zwischen den eingegebenen Stützpunkten stückweise linear interpoliert. Keine Polynom-Anpassung, Glättung oder Extrapolation. Herstellerdaten für das tatsächliche Medium verwenden; Drehzahl-, Viskositäts- und Dichtekorrekturen außerhalb der einfachen Förderhöhen-Druck-Umrechnung erfolgen nicht automatisch.

Gesucht wird:

```text
F(Q) = Δp_Pumpe(Q) − Δp_System(Q) = 0
```

Bei fallender Pumpen- und steigender Systemkennlinie gibt es höchstens einen Schnittpunkt. Die Druckdifferenzen an den beiden Tabellenenden werden geprüft; ein Randpunkt mit |F|≤10⁻⁸ bar wird akzeptiert. Andernfalls muss F links positiv und rechts negativ sein. Danach wird mit Bisektion über höchstens 120 Schritte gelöst; nur ein tatsächlich kleiner Druckfehler |F|≤10⁻⁸ bar gilt als Erfolg. Eine kleine Q-Intervallbreite allein reicht nicht aus, da das gewählte Übergangsmodell einen Sprung aufweisen kann. Gibt es an diesem Sprung keinen echten Schnittpunkt, wird kein Betriebspunkt erfunden.

Die Darstellung der Systemkurve im Pumpendiagramm verwendet 201 Punkte über den Pumpentabellenbereich. Der Schnittpunkt wird unabhängig von dieser grafischen Abtastung berechnet und zusätzlich markiert. Ergebnisse und Warnungen des Betriebspunkts stehen auf `Betriebspunkte`; das Druckresiduum steht in O11. Der Kennlinienbereich der normalen Systemkurve hat keinen Einfluss auf die Pumpensuche.

Eine lokale Steigungsbedingung mit fallender Pumpenkennlinie ersetzt keine vollständige dynamische Stabilitätsprüfung. Wirkungsgrad, Leistungsbedarf, NPSH, Mindestförderstrom, Parallel-/Serienbetrieb und zulässiger Hersteller-Betriebsbereich sind nicht Bestandteil dieser Version. Ein rechnerischer Schnittpunkt ist daher noch keine Freigabe der Pumpenauslegung.
