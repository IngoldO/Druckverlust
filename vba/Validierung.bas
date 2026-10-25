Attribute VB_Name = "Validierung"
Option Explicit

Private Sub Nah(ByVal actual As Double, ByVal expected As Double, ByVal label As String)
    If Abs(actual - expected) > 0.00000001 * (1 + Abs(expected)) Then Err.Raise vbObjectError + 100, , label & ": Ist=" & actual & " Soll=" & expected
End Sub

Public Sub ValidierungStarten()
    On Error GoTo Fehler
    Dim r As Betriebspunkt, count As Long, i As Long, bad As Variant, failures As Long, code As Long
    r = BerechnePunkt(0.10000000000000001#, 0.050000000000000003#, 10#, 0#, 1000#, 0.01#, 0#, 0#, 0#)
    Nah r.V, 0.014147106052612919#, "Laminar: V"
    Nah r.Re, 70.735530263064604#, "Laminar: Re"
    Nah r.Lambda, 0.90477868423386032#, "Laminar: Lambda"
    Nah r.RohrBar, 0.00018108295747344534#, "Laminar: RohrBar"
    Nah r.EinzelBar, 0#, "Laminar: EinzelBar"
    Nah r.HoeheBar, 0#, "Laminar: HoeheBar"
    Nah r.BehaelterBar, 0#, "Laminar: BehaelterBar"
    Nah r.GesamtBar, 0.00018108295747344534#, "Laminar: GesamtBar"
    count = count + 1
    r = BerechnePunkt(10#, 0.050000000000000003#, 100#, 0#, 998.20000000000005#, 0.0010020000000000001#, 0#, 0#, 0#)
    Nah r.V, 1.4147106052612919#, "Turbulent glatt: V"
    Nah r.Re, 70467.271765060956#, "Turbulent glatt: Re"
    Nah r.Lambda, 0.019376650122451919#, "Turbulent glatt: Lambda"
    Nah r.RohrBar, 0.38710740705274138#, "Turbulent glatt: RohrBar"
    Nah r.EinzelBar, 0#, "Turbulent glatt: EinzelBar"
    Nah r.HoeheBar, 0#, "Turbulent glatt: HoeheBar"
    Nah r.BehaelterBar, 0#, "Turbulent glatt: BehaelterBar"
    Nah r.GesamtBar, 0.38710740705274138#, "Turbulent glatt: GesamtBar"
    count = count + 1
    r = BerechnePunkt(30#, 0.050000000000000003#, 100#, 0.0001#, 1000#, 0.001#, 0#, 0#, 0#)
    Nah r.V, 4.2441318157838754#, "Turbulent rau: V"
    Nah r.Re, 212206.59078919375#, "Turbulent rau: Re"
    Nah r.Lambda, 0.02426095017210908#, "Turbulent rau: Lambda"
    Nah r.RohrBar, 4.3700412226237679#, "Turbulent rau: RohrBar"
    Nah r.EinzelBar, 0#, "Turbulent rau: EinzelBar"
    Nah r.HoeheBar, 0#, "Turbulent rau: HoeheBar"
    Nah r.BehaelterBar, 0#, "Turbulent rau: BehaelterBar"
    Nah r.GesamtBar, 4.3700412226237679#, "Turbulent rau: GesamtBar"
    count = count + 1
    r = BerechnePunkt(10#, 0.050000000000000003#, 100#, 4.5000000000000003E-05, 1000#, 0.001#, 0#, 13#, 0#)
    Nah r.V, 1.4147106052612919#, "Mehrere Armaturen: V"
    Nah r.Re, 70735.530263064604#, "Mehrere Armaturen: Re"
    Nah r.Lambda, 0.022683750787681402#, "Mehrere Armaturen: Lambda"
    Nah r.RohrBar, 0.45399397121100077#, "Mehrere Armaturen: RohrBar"
    Nah r.EinzelBar, 0.13009139628152011#, "Mehrere Armaturen: EinzelBar"
    Nah r.HoeheBar, 0#, "Mehrere Armaturen: HoeheBar"
    Nah r.BehaelterBar, 0#, "Mehrere Armaturen: BehaelterBar"
    Nah r.GesamtBar, 0.58408536749252082#, "Mehrere Armaturen: GesamtBar"
    count = count + 1
    r = BerechnePunkt(10#, 0.050000000000000003#, 100#, 0#, 1000#, 0.001#, 12#, 0#, 0#)
    Nah r.V, 1.4147106052612919#, "Geodaetisch: V"
    Nah r.Re, 70735.530263064604#, "Geodaetisch: Re"
    Nah r.Lambda, 0.019360776505972434#, "Geodaetisch: Lambda"
    Nah r.RohrBar, 0.38748776134713914#, "Geodaetisch: RohrBar"
    Nah r.EinzelBar, 0#, "Geodaetisch: EinzelBar"
    Nah r.HoeheBar, 1.1767979999999998#, "Geodaetisch: HoeheBar"
    Nah r.BehaelterBar, 0#, "Geodaetisch: BehaelterBar"
    Nah r.GesamtBar, 1.564285761347139#, "Geodaetisch: GesamtBar"
    count = count + 1
    r = BerechnePunkt(0#, 0.050000000000000003#, 100#, 4.5000000000000003E-05, 1000#, 0.001#, 12#, 3#, 1.5#)
    Nah r.V, 0#, "Stillstand: V"
    Nah r.Re, 0#, "Stillstand: Re"
    Nah r.Lambda, 0#, "Stillstand: Lambda"
    Nah r.RohrBar, 0#, "Stillstand: RohrBar"
    Nah r.EinzelBar, 0#, "Stillstand: EinzelBar"
    Nah r.HoeheBar, 1.1767979999999998#, "Stillstand: HoeheBar"
    Nah r.BehaelterBar, 1.5#, "Stillstand: BehaelterBar"
    Nah r.GesamtBar, 2.6767979999999998#, "Stillstand: GesamtBar"
    count = count + 1
    r = BerechnePunkt(0#, 0.050000000000000003#, 0#, 0#, 1000#, 0.001#, -10#, 0#, -0.5#)
    Nah r.V, 0#, "Negativer statischer Anteil: V"
    Nah r.Re, 0#, "Negativer statischer Anteil: Re"
    Nah r.Lambda, 0#, "Negativer statischer Anteil: Lambda"
    Nah r.RohrBar, 0#, "Negativer statischer Anteil: RohrBar"
    Nah r.EinzelBar, 0#, "Negativer statischer Anteil: EinzelBar"
    Nah r.HoeheBar, -0.98066500000000001#, "Negativer statischer Anteil: HoeheBar"
    Nah r.BehaelterBar, -0.5#, "Negativer statischer Anteil: BehaelterBar"
    Nah r.GesamtBar, -1.4806650000000001#, "Negativer statischer Anteil: GesamtBar"
    count = count + 1
    r = BerechnePunkt(0.42411500823462206#, 0.050000000000000003#, 10#, 0#, 1000#, 0.001#, 0#, 0#, 0#)
    Nah r.V, 0.059999999999999991#, "Uebergang: V"
    Nah r.Re, 3000#, "Uebergang: Re"
    Nah r.Lambda, 0.0435191887685763#, "Uebergang: Lambda"
    Nah r.RohrBar, 0.00015666907956687464#, "Uebergang: RohrBar"
    Nah r.EinzelBar, 0#, "Uebergang: EinzelBar"
    Nah r.HoeheBar, 0#, "Uebergang: HoeheBar"
    Nah r.BehaelterBar, 0#, "Uebergang: BehaelterBar"
    Nah r.GesamtBar, 0.00015666907956687464#, "Uebergang: GesamtBar"
    count = count + 1
    r = BerechnePunkt(0.32515483964654363#, 0.050000000000000003#, 10#, 0#, 1000#, 0.001#, 0#, 0#, 0#)
    Nah r.V, 0.046000000000000006#, "Re2300: V"
    Nah r.Re, 2300#, "Re2300: Re"
    Nah r.Lambda, 0.047283313905224833#, "Re2300: Lambda"
    Nah r.RohrBar, 0.00010005149222345576#, "Re2300: RohrBar"
    Nah r.EinzelBar, 0#, "Re2300: EinzelBar"
    Nah r.HoeheBar, 0#, "Re2300: HoeheBar"
    Nah r.BehaelterBar, 0#, "Re2300: BehaelterBar"
    Nah r.GesamtBar, 0.00010005149222345576#, "Re2300: GesamtBar"
    count = count + 1
    r = BerechnePunkt(0.56548667764616278#, 0.050000000000000003#, 10#, 0#, 1000#, 0.001#, 0#, 0#, 0#)
    Nah r.V, 0.079999999999999988#, "Re4000: V"
    Nah r.Re, 3999.9999999999995#, "Re4000: Re"
    Nah r.Lambda, 0.03990701405563489#, "Re4000: Lambda"
    Nah r.RohrBar, 0.00025540488995606325#, "Re4000: RohrBar"
    Nah r.EinzelBar, 0#, "Re4000: EinzelBar"
    Nah r.HoeheBar, 0#, "Re4000: HoeheBar"
    Nah r.BehaelterBar, 0#, "Re4000: BehaelterBar"
    Nah r.GesamtBar, 0.00025540488995606325#, "Re4000: GesamtBar"
    count = count + 1
    r = BerechnePunkt(10#, 0.050000000000000003#, 0#, 0#, 1000#, 0.001#, 0#, 5#, 0#)
    Nah r.V, 1.4147106052612919#, "Nur Einzelwiderstaende: V"
    Nah r.Re, 70735.530263064604#, "Nur Einzelwiderstaende: Re"
    Nah r.Lambda, 0.019360776505972434#, "Nur Einzelwiderstaende: Lambda"
    Nah r.RohrBar, 0#, "Nur Einzelwiderstaende: RohrBar"
    Nah r.EinzelBar, 0.050035152415969272#, "Nur Einzelwiderstaende: EinzelBar"
    Nah r.HoeheBar, 0#, "Nur Einzelwiderstaende: HoeheBar"
    Nah r.BehaelterBar, 0#, "Nur Einzelwiderstaende: BehaelterBar"
    Nah r.GesamtBar, 0.050035152415969272#, "Nur Einzelwiderstaende: GesamtBar"
    count = count + 1
    r = BerechnePunkt(10#, 0.050000000000000003#, 10#, 0.0025000000000000001#, 1000#, 0.001#, 0#, 0#, 0#)
    Nah r.V, 1.4147106052612919#, "Maximale Rauigkeit: V"
    Nah r.Re, 70735.530263064604#, "Maximale Rauigkeit: Re"
    Nah r.Lambda, 0.071875835562408591#, "Maximale Rauigkeit: Lambda"
    Nah r.RohrBar, 0.14385273549561034#, "Maximale Rauigkeit: RohrBar"
    Nah r.EinzelBar, 0#, "Maximale Rauigkeit: EinzelBar"
    Nah r.HoeheBar, 0#, "Maximale Rauigkeit: HoeheBar"
    Nah r.BehaelterBar, 0#, "Maximale Rauigkeit: BehaelterBar"
    Nah r.GesamtBar, 0.14385273549561034#, "Maximale Rauigkeit: GesamtBar"
    count = count + 1
    For i = 1 To 8
        On Error Resume Next
        Err.Clear
        Select Case i
        Case 1: r = BerechnePunkt(-1, 0.05, 10, 0, 1000, 0.001, 0, 0, 0)
        Case 2: r = BerechnePunkt(1, 0, 10, 0, 1000, 0.001, 0, 0, 0)
        Case 3: r = BerechnePunkt(1, 0.05, -1, 0, 1000, 0.001, 0, 0, 0)
        Case 4: r = BerechnePunkt(1, 0.05, 10, -1, 1000, 0.001, 0, 0, 0)
        Case 5: r = BerechnePunkt(1, 0.05, 10, 0, 0, 0.001, 0, 0, 0)
        Case 6: r = BerechnePunkt(1, 0.05, 10, 0, 1000, 0, 0, 0, 0)
        Case 7: r = BerechnePunkt(1, 0.05, 10, 0, 1000, 0.001, 0, -1, 0)
        Case 8: r = BerechnePunkt(1, 0.05, 10, 0.003, 1000, 0.001, 0, 0, 0)
        End Select
        code = Err.Number
        On Error GoTo Fehler
        If code = 0 Then Err.Raise vbObjectError + 101, , "Ungueltige Eingabe akzeptiert, Fall " & i
        count = count + 1
    Next i
    MsgBox count & " Rechenkern-Testfaelle bestanden (12 Referenzen, 8 Fehlerfaelle).", vbInformation
    Exit Sub
Fehler:
    MsgBox "VALIDIERUNG FEHLGESCHLAGEN: " & Err.Description, vbCritical
    Err.Raise vbObjectError + 102, "Validierung", "Validierung fehlgeschlagen"
End Sub
