Attribute VB_Name = "Hydraulik"
Option Explicit

Public Const G As Double = 9.80665
Public Type Betriebspunkt
    Q As Double
    V As Double
    Re As Double
    Lambda As Double
    RohrBar As Double
    EinzelBar As Double
    HoeheBar As Double
    BehaelterBar As Double
    GesamtBar As Double
    Regime As String
    Hinweis As String
End Type

Public Type Anlagendaten
    D As Double
    L As Double
    Eps As Double
    Rho As Double
    Mu As Double
    H As Double
    DP As Double
    Z As Double
    ZKv As Double
    ZBlende As Double
    BlendeD As Double
End Type

' Darcy factor; Colebrook fixed point stops at absolute change <1E-10 in 1/sqrt(lambda).
' Transition: upper envelope of laminar and turbulent branches, NOT interpolation.
Public Function Reibungsbeiwert(ByVal re As Double, ByVal relRau As Double) As Double
    Dim x As Double, xn As Double, i As Long
    If re < 0 Or relRau < 0 Or relRau > 0.05 Then Err.Raise vbObjectError + 1, , "Re oder relative Rauigkeit ausserhalb Modellbereich."
    If re = 0 Then Exit Function
    If re < 2300 Then
        Reibungsbeiwert = 64 / re
        Exit Function
    End If
    x = 7
    For i = 1 To 200
        xn = -2 * Log(relRau / 3.7 + 2.51 * x / re) / Log(10#)
        If Abs(xn - x) < 0.0000000001 Then
            Reibungsbeiwert = 1 / (xn * xn)
            If re < 4000 And Reibungsbeiwert < 64 / re Then Reibungsbeiwert = 64 / re
            Exit Function
        End If
        x = xn
    Next i
    Err.Raise vbObjectError + 2, , "Colebrook-Verfahren konvergiert nicht."
End Function

Public Function BerechnePunkt(ByVal q As Double, ByVal d As Double, ByVal l As Double, ByVal eps As Double, ByVal rho As Double, ByVal mu As Double, ByVal h As Double, ByVal z As Double, ByVal dp As Double) As Betriebspunkt
    Dim r As Betriebspunkt, dyn As Double
    ' SI: Q m3/h, D/L/eps/h m, rho kg/m3, mu Pa s, dp bar.
    If q < 0 Or d <= 0 Or l < 0 Or eps < 0 Or rho <= 0 Or mu <= 0 Or z < 0 Then Err.Raise vbObjectError + 3, , "Unzulaessige Eingaben im Rechenkern."
    If eps / d > 0.05 Then Err.Raise vbObjectError + 4, , "Relative Rauigkeit muss <= 0.05 sein."
    r.q = q
    r.V = (q / 3600#) / (4 * Atn(1#) * d * d / 4)
    r.re = rho * r.V * d / mu
    r.Lambda = Reibungsbeiwert(r.re, eps / d)
    dyn = rho * r.V * r.V / 2 / 100000#
    r.RohrBar = r.Lambda * (l / d) * dyn
    r.EinzelBar = z * dyn
    r.HoeheBar = rho * G * h / 100000#
    r.BehaelterBar = dp
    r.GesamtBar = r.RohrBar + r.EinzelBar + r.HoeheBar + dp
    If q = 0 Then
        r.Regime = "Stillstand"
        r.Hinweis = "Lambda=0 als Rechenkonvention; nur statische Druckanteile."
    ElseIf r.re < 2300 Then
        r.Regime = "Laminar"
    ElseIf r.re < 4000 Then
        r.Regime = "Uebergang"
        r.Hinweis = "UNSICHER: Maximum aus 64/Re und Colebrook; keine gesicherte Genauigkeit."
    Else
        r.Regime = "Turbulent"
    End If
    If z > 0 And r.re > 0 And r.re < 4000 Then r.Hinweis = r.Hinweis & " Konstante Zeta-Werte bei kleinem Re besonders unsicher."
    If r.V > 5 Then r.Hinweis = r.Hinweis & " Hohe Geschwindigkeit >5 m/s: Auslegung pruefen."
    If r.GesamtBar < 0 Then r.Hinweis = r.Hinweis & " Negative Druckerhoehung: statischer Antrieb; Systemgrenzen pruefen."
    BerechnePunkt = r
End Function

' Permanent loss of a thin sharp-edged concentric orifice, not tapping differential.
' Cd is an explicitly assumed constant; no ISO 5167 metering certification.
Public Function BlendenZeta(ByVal rohrD As Double, ByVal lochD As Double, ByVal cd As Double) As Double
    Dim beta As Double, a As Double
    If rohrD <= 0 Or lochD <= 0 Or cd <= 0 Or cd >= 1 Then Err.Raise vbObjectError + 30, , "Blende: positive Durchmesser und 0 < Cd < 1 erforderlich."
    beta = lochD / rohrD
    If beta < 0.2 - 0.000000000001 Or beta > 0.75 + 0.000000000001 Then Err.Raise vbObjectError + 31, , "Blende: Modellbereich 0.2 <= d/D <= 0.75."
    If beta < 0.2 Then beta = 0.2
    If beta > 0.75 Then beta = 0.75
    a = Sqr(1 - beta ^ 4 * (1 - cd ^ 2))
    BlendenZeta = (a / (cd * beta ^ 2) - 1) ^ 2
End Function

' Kv [m3/h] at 1 bar with reference water density 1000 kg/m3.
Public Function KvDruck(ByVal q As Double, ByVal kv As Double, ByVal rho As Double) As Double
    If q < 0 Or kv <= 0 Or rho <= 0 Then Err.Raise vbObjectError + 60, , "Kv: Q >= 0, Kv > 0 und Dichte > 0 erforderlich."
    KvDruck = (rho / 1000#) * (q / kv) ^ 2
End Function
