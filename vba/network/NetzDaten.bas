Attribute VB_Name = "NetzDaten"
Option Explicit

Public Type NetzKnoten
    ID As String
    Bezeichnung As String
    Typ As String
    Hoehe As Double
    DruckFest As Boolean
    DruckBarA As Double
    FlussVorgegeben As Boolean
    ExternQ As Double ' m3/h: positive injection, negative withdrawal
    X As Double
    Y As Double
End Type

Public Type NetzLeitung
    ID As String
    Bezeichnung As String
    Von As Long
    Nach As Long
    DN As Double
    D As Double ' m
    L As Double ' m
    Eps As Double ' m
    Zeta As Double ' ordinary fittings ONLY
    ZKv As Double
    ZBlende As Double
    Aktiv As Boolean
    Pumpe As Boolean
    Rueckschlag As Boolean ' oriented from Von to Nach
End Type

Public Type NetzResultat
    Erfolg As Boolean
    Status As String
    Iterationen As Long
    MaxBilanz As Double ' m3/h, free nodes ONLY
    MaxMasse As Double ' kg/s, free nodes ONLY
    MaxEnergie As Double ' bar
    PivotVerhaeltnis As Double ' diagnostic, NOT a norm condition number
    Druck() As Double ' bar(a)
    Q() As Double ' m3/h, signed in reference direction; inactive = 0
    Bilanz() As Double ' external reaction required at fixed-pressure nodes
    Energie() As Double
End Type

Public Sub NetzFehler(ByVal txt As String)
    Err.Raise vbObjectError + 200, "Netzwerk", txt
End Sub

Public Function NetzVerlust(ByVal q As Double, ByRef e As NetzLeitung, ByVal rho As Double, ByVal mu As Double) As Double
    Dim r As Betriebspunkt
    r = BerechnePunkt(Abs(q), e.D, e.L, e.Eps, rho, mu, 0, e.Zeta + e.ZKv + e.ZBlende, 0)
    NetzVerlust = Sgn(q) * (r.RohrBar + r.EinzelBar)
End Function

Public Function NetzKvZeta(ByVal kv As Double, ByVal n As Double, ByVal d As Double, ByVal rho As Double) As Double
    If n < 0 Or n <> Fix(n) Or d <= 0 Then NetzFehler "Kv: Anzahl ganzzahlig >=0 und D>0 erforderlich."
    NetzKvZeta = n * KvDruck(1, kv, rho) * 200000# / rho * (3600# * 4 * Atn(1#) * d ^ 2 / 4) ^ 2
End Function

Public Function NetzBlendenZeta(ByVal lokD As Double, ByVal lochD As Double, ByVal cd As Double, ByVal n As Double, ByVal d As Double) As Double
    If n < 0 Or n <> Fix(n) Or d <= 0 Then NetzFehler "Blende: Anzahl ganzzahlig >=0 und D>0 erforderlich."
    NetzBlendenZeta = n * BlendenZeta(lokD, lochD, cd) * (d / lokD) ^ 4
End Function

Public Sub NetzPruefen(ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByVal rho As Double, ByVal mu As Double)
    Dim i As Long, j As Long, n As Long, fest As Long, pumpen As Long, a As Long, b As Long
    Dim seen() As Boolean, changed As Boolean, r As Betriebspunkt
    If LBound(k) <> 1 Or LBound(e) <> 1 Then NetzFehler "Netzarrays muessen bei 1 beginnen."
    n = UBound(k)
    If n < 2 Or rho <= 0 Or mu <= 0 Then NetzFehler "Mindestens 2 Knoten und positive Stoffwerte erforderlich."
    ReDim seen(1 To n)
    For i = 1 To n
        If Len(Trim(k(i).ID)) = 0 Then NetzFehler "Knoten-ID fehlt."
        For j = 1 To i - 1
            If StrComp(k(i).ID, k(j).ID, vbTextCompare) = 0 Then NetzFehler "Doppelte Knoten-ID: " & k(i).ID
        Next j
        If k(i).DruckFest Then
            fest = fest + 1
            If k(i).DruckBarA <= 0 Then NetzFehler "Absolutdruck muss positiv sein: " & k(i).ID
            If k(i).FlussVorgegeben Or k(i).ExternQ <> 0 Then NetzFehler "Ueberbestimmt: Druck und externer Fluss am selben Knoten."
        ElseIf k(i).ExternQ <> 0 And Not k(i).FlussVorgegeben Then
            NetzFehler "Externer Fluss ohne Flussrandbedingung: " & k(i).ID
        End If
    Next i
    If fest = 0 Then NetzFehler "Druckrandbedingung fehlt; mindestens ein positiver Absolutdruck erforderlich."
    For i = 1 To UBound(e)
        If Len(Trim(e(i).ID)) = 0 Then NetzFehler "Leitungs-ID fehlt."
        For j = 1 To i - 1
            If StrComp(e(i).ID, e(j).ID, vbTextCompare) = 0 Then NetzFehler "Doppelte Leitungs-ID: " & e(i).ID
        Next j
        If e(i).Aktiv Then
            a = e(i).Von: b = e(i).Nach
            If a < 1 Or a > n Or b < 1 Or b > n Or a = b Then NetzFehler "Ungueltige Knotenreferenz: " & e(i).ID
            r = BerechnePunkt(0, e(i).D, e(i).L, e(i).Eps, rho, mu, 0, e(i).Zeta, 0)
            If e(i).ZKv < 0 Or e(i).ZBlende < 0 Then NetzFehler "Negative Teil-Verlustbeiwerte."
            If e(i).Pumpe Then pumpen = pumpen + 1
            If e(i).L = 0 And e(i).Zeta + e(i).ZKv + e(i).ZBlende = 0 And Not e(i).Pumpe Then NetzFehler "Widerstandslose Leitung: " & e(i).ID & ". Knoten zusammenlegen."
        End If
    Next i
    If pumpen > 1 Then NetzFehler "Diese Version unterstuetzt hoechstens eine aktive Pumpe."
    seen(1) = True
    Do
        changed = False
        For i = 1 To UBound(e)
            If e(i).Aktiv Then
                a = e(i).Von: b = e(i).Nach
                If seen(a) Xor seen(b) Then
                    seen(a) = True: seen(b) = True: changed = True
                End If
            End If
        Next i
    Loop While changed
    For i = 1 To n
        If Not seen(i) Then NetzFehler "Nicht verbundenes Teilnetz / isolierter Knoten: " & k(i).ID
    Next i
End Sub
