Attribute VB_Name = "NetzSolver"
Option Explicit

Private Function Maximum(ByRef f() As Double) As Double
    Dim i As Long
    For i = 1 To UBound(f)
        If Abs(f(i)) > Maximum Then Maximum = Abs(f(i))
    Next i
End Function

' Partial pivoting on a row-scaled Jacobian, normalized unknowns.
Private Sub Gauss(ByRef a() As Double, ByRef b() As Double, ByRef dx() As Double, ByRef ratio As Double)
    Dim n As Long, i As Long, j As Long, k As Long, pivot As Long
    Dim scale As Double, v As Double, t As Double, big As Double, small As Double
    n = UBound(b): ReDim dx(1 To n)
    For i = 1 To n
        scale = 0
        For j = 1 To n
            If Abs(a(i, j)) > scale Then scale = Abs(a(i, j))
        Next j
        If scale = 0 Then NetzFehler "Singulaere Jacobi-Matrix: unzureichende Randbedingungen."
        For j = 1 To n
            a(i, j) = a(i, j) / scale
        Next j
        b(i) = b(i) / scale
    Next i
    small = 1E+100
    For k = 1 To n
        pivot = k: v = Abs(a(k, k))
        For i = k + 1 To n
            If Abs(a(i, k)) > v Then pivot = i: v = Abs(a(i, k))
        Next i
        If v < 0.000000000001 Then NetzFehler "Singulaeres oder schlecht konditioniertes Netz; Randbedingungen/Widerstaende pruefen."
        If v > big Then big = v
        If v < small Then small = v
        If pivot <> k Then
            For j = k To n
                t = a(k, j): a(k, j) = a(pivot, j): a(pivot, j) = t
            Next j
            t = b(k): b(k) = b(pivot): b(pivot) = t
        End If
        For i = k + 1 To n
            t = a(i, k) / a(k, k)
            For j = k + 1 To n
                a(i, j) = a(i, j) - t * a(k, j)
            Next j
            b(i) = b(i) - t * b(k)
        Next i
    Next k
    ratio = big / small
    For i = n To 1 Step -1
        t = b(i)
        For j = i + 1 To n
            t = t - a(i, j) * dx(j)
        Next j
        dx(i) = t / a(i, i)
    Next i
End Sub

Private Sub Residuum(ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByRef im() As Long, ByRef ip() As Long, ByVal nf As Long, ByVal m As Long, ByVal festQ As Double, ByRef x() As Double, ByVal qs As Double, ByVal ps As Double, ByVal rho As Double, ByVal mu As Double, ByRef pq() As Double, ByRef ph() As Double, ByVal tq As Double, ByVal tp As Double, ByRef f() As Double)
    Dim i As Long, a As Long, b As Long, q As Double, drive As Double, loss As Double
    Dim p() As Double, bal() As Double, eq As Double
    ReDim f(1 To UBound(x)): ReDim p(1 To UBound(k)): ReDim bal(1 To UBound(k))
    For i = 1 To UBound(k)
        If ip(i) = 0 Then p(i) = k(i).DruckBarA Else p(i) = x(ip(i)) * ps
        bal(i) = k(i).ExternQ
    Next i
    For i = 1 To UBound(e)
        If im(i) > 0 Then
            a = e(i).Von: b = e(i).Nach: q = x(im(i)) * qs
            bal(a) = bal(a) - q: bal(b) = bal(b) + q
            drive = p(a) - p(b) + rho * G * (k(a).Hoehe - k(b).Hoehe) / 100000#
            If e(i).Pumpe Then
                If festQ >= 0 Then drive = drive + x(UBound(x)) * ps Else drive = drive + Pumpendruck(q, pq, ph, rho)
                If festQ >= 0 Then f(UBound(x)) = (q - festQ) / tq
            End If
            loss = NetzVerlust(q, e(i), rho, mu)
            eq = (loss - drive) / tp
            If e(i).Rueckschlag Then
                ' Complementarity: q>=0, loss-drive>=0, q*(loss-drive)=0.
                If q / tq < eq Then eq = q / tq
            End If
            f(nf + im(i)) = eq
        End If
    Next i
    For i = 1 To UBound(k)
        If ip(i) > 0 Then f(ip(i) - m) = bal(i) / tq
    Next i
End Sub

Public Sub NetzLoesen(ByRef k() As NetzKnoten, ByRef e() As NetzLeitung, ByVal rho As Double, ByVal mu As Double, ByRef pq() As Double, ByRef ph() As Double, ByRef aus As NetzResultat, Optional ByVal maxIter As Long = 100, Optional ByVal tolQ As Double = 0.000001, Optional ByVal tolP As Double = 0.00000001, Optional ByVal festPumpQ As Double = -1)
    On Error GoTo Fehler
    Dim im() As Long, ip() As Long, x() As Double, trial() As Double, f() As Double, ft() As Double
    Dim jac() As Double, rhs() As Double, dx() As Double, tmp As NetzResultat
    Dim m As Long, nf As Long, dimn As Long, i As Long, j As Long, it As Long, back As Long, pump As Long
    Dim qs As Double, ps As Double, meanHead As Double, count As Long, step As Double, old As Double
    Dim merit As Double, alpha As Double, ratio As Double, feasible As Boolean, accepted As Boolean
    Dim q As Double, drive As Double, energy As Double, message As String
    Dim connected() As Boolean, changed As Boolean
    aus = tmp: aus.Status = "NICHT GELOEST"
    If maxIter < 1 Or tolQ <= 0 Or tolP <= 0 Then NetzFehler "Iterationslimit/Toleranzen ungueltig."
    NetzPruefen k, e, rho, mu
    ReDim im(1 To UBound(e)): ReDim ip(1 To UBound(k))
    qs = 1: ps = 1
    For i = 1 To UBound(e)
        If e(i).Aktiv Then
            m = m + 1: im(i) = m
            If e(i).Pumpe Then pump = i
        End If
    Next i
    For i = 1 To UBound(k)
        If Abs(k(i).ExternQ) > qs Then qs = Abs(k(i).ExternQ)
        If k(i).DruckFest Then
            count = count + 1
            meanHead = meanHead + k(i).DruckBarA + rho * G * k(i).Hoehe / 100000#
            If k(i).DruckBarA > ps Then ps = k(i).DruckBarA
        Else
            nf = nf + 1: ip(i) = m + nf
        End If
    Next i
    If festPumpQ < 0 And festPumpQ <> -1 Then NetzFehler "Pumpen-Q-Vorgabe muss >=0 oder -1 (Betriebspunkt) sein."
    If festPumpQ >= 0 And pump = 0 Then NetzFehler "Pumpen-Q-Vorgabe ohne aktive Pumpe."
    meanHead = meanHead / count: dimn = m + nf
    If festPumpQ >= 0 Then dimn = dimn + 1
    ReDim x(1 To dimn): ReDim trial(1 To dimn)
    If pump > 0 And festPumpQ < 0 Then
        PruefePumpendaten pq, ph, rho
        If pq(LBound(pq)) <> 0 Then NetzFehler "Netzpumpe: Kennlinie muss bei Q=0 beginnen; keine Extrapolation."
    End If
    For i = 1 To UBound(e)
        If im(i) > 0 Then x(im(i)) = 0.1
    Next i
    If pump > 0 Then
        If festPumpQ >= 0 Then
            x(im(pump)) = festPumpQ / qs: x(dimn) = 1
        Else
            x(im(pump)) = (pq(LBound(pq)) + pq(UBound(pq))) / (2 * qs)
        End If
    End If
    For i = 1 To UBound(k)
        If ip(i) > 0 Then x(ip(i)) = (meanHead - rho * G * k(i).Hoehe / 100000#) / ps
    Next i
    For it = 0 To maxIter
        Residuum k, e, im, ip, nf, m, festPumpQ, x, qs, ps, rho, mu, pq, ph, tolQ, tolP, f
        merit = Maximum(f)
        If merit <= 1 Then Exit For
        If it = maxIter Then NetzFehler "Keine Konvergenz nach " & maxIter & " Iterationen. Keine gueltigen Resultate."
        ReDim jac(1 To dimn, 1 To dimn): ReDim rhs(1 To dimn)
        For j = 1 To dimn
            old = x(j): step = 0.000001 * (1 + Abs(old))
            If pump > 0 And festPumpQ < 0 Then
                If j = im(pump) And (old + step) * qs > pq(UBound(pq)) Then step = -step
                If j = im(pump) And (old + step) * qs < 0 Then step = (pq(UBound(pq)) / qs - old) / 2
            End If
            x(j) = old + step
            Residuum k, e, im, ip, nf, m, festPumpQ, x, qs, ps, rho, mu, pq, ph, tolQ, tolP, ft
            x(j) = old
            For i = 1 To dimn
                jac(i, j) = (ft(i) - f(i)) / step
            Next i
        Next j
        For i = 1 To dimn
            rhs(i) = -f(i)
        Next i
        Gauss jac, rhs, dx, ratio
        alpha = 1: accepted = False
        For back = 1 To 40
            For i = 1 To dimn
                trial(i) = x(i) + alpha * dx(i)
            Next i
            feasible = True
            If pump > 0 And festPumpQ < 0 Then
                q = trial(im(pump)) * qs
                feasible = (q >= 0 And q <= pq(UBound(pq)))
            End If
            If feasible Then
                Residuum k, e, im, ip, nf, m, festPumpQ, trial, qs, ps, rho, mu, pq, ph, tolQ, tolP, ft
                If Maximum(ft) < merit * (1 - 0.0001 * alpha) Or Maximum(ft) <= 1 Then
                    accepted = True: Exit For
                End If
            End If
            alpha = alpha / 2
        Next back
        If Not accepted Then NetzFehler "Newton-Daempfung ohne Fortschritt: Randbedingungen, Modellwechsel oder Pumpenbereich pruefen."
        For i = 1 To dimn
            x(i) = trial(i)
        Next i
    Next it
    ReDim tmp.Druck(1 To UBound(k)): ReDim tmp.Bilanz(1 To UBound(k))
    ReDim tmp.Q(1 To UBound(e)): ReDim tmp.Energie(1 To UBound(e))
    For i = 1 To UBound(k)
        If ip(i) = 0 Then tmp.Druck(i) = k(i).DruckBarA Else tmp.Druck(i) = x(ip(i)) * ps
        If tmp.Druck(i) <= 0 Then NetzFehler "Unphysikalischer Absolutdruck <=0 am Knoten " & k(i).ID
        tmp.Bilanz(i) = k(i).ExternQ
    Next i
    For i = 1 To UBound(e)
        If im(i) > 0 Then
            q = x(im(i)) * qs: tmp.Q(i) = q
            tmp.Bilanz(e(i).Von) = tmp.Bilanz(e(i).Von) - q
            tmp.Bilanz(e(i).Nach) = tmp.Bilanz(e(i).Nach) + q
            drive = tmp.Druck(e(i).Von) - tmp.Druck(e(i).Nach) + rho * G * (k(e(i).Von).Hoehe - k(e(i).Nach).Hoehe) / 100000#
            If e(i).Pumpe Then
                If festPumpQ >= 0 Then tmp.PumpenDruck = x(dimn) * ps Else tmp.PumpenDruck = Pumpendruck(q, pq, ph, rho)
                tmp.PumpenQ = q
                drive = drive + tmp.PumpenDruck
            End If
            energy = NetzVerlust(q, e(i), rho, mu) - drive
            If e(i).Rueckschlag And Abs(q) <= tolQ Then
                If energy > 0 Then energy = 0 ' closed valve supports reverse pressure
            End If
            tmp.Energie(i) = energy
            If Abs(energy) > tmp.MaxEnergie Then tmp.MaxEnergie = Abs(energy)
        End If
    Next i
    For i = 1 To UBound(k)
        If ip(i) > 0 And Abs(tmp.Bilanz(i)) > tmp.MaxBilanz Then tmp.MaxBilanz = Abs(tmp.Bilanz(i))
    Next i
    ' A shut check valve can isolate free pressures even in a connected input graph.
    ReDim connected(1 To UBound(k))
    For i = 1 To UBound(k)
        connected(i) = k(i).DruckFest
    Next i
    Do
        changed = False
        For i = 1 To UBound(e)
            If e(i).Aktiv Then
                If Not (e(i).Rueckschlag And Abs(tmp.Q(i)) <= tolQ) Then
                    If connected(e(i).Von) Xor connected(e(i).Nach) Then
                        connected(e(i).Von) = True: connected(e(i).Nach) = True: changed = True
                    End If
                End If
            End If
        Next i
    Loop While changed
    For i = 1 To UBound(k)
        If Not connected(i) Then NetzFehler "Nicht eindeutig bestimmter Druck hinter geschlossenen Rueckschlagventilen: " & k(i).ID
    Next i
    tmp.MaxMasse = tmp.MaxBilanz * rho / 3600#
    If tmp.MaxBilanz > tolQ Or tmp.MaxEnergie > tolP Then NetzFehler "Abschliessende Bilanzpruefung fehlgeschlagen."
    tmp.Erfolg = True: tmp.Status = "KONVERGIERT": tmp.Iterationen = it: tmp.PivotVerhaeltnis = ratio
    aus = tmp
    Exit Sub
Fehler:
    message = Err.Description
    Dim leer As NetzResultat
    aus = leer: aus.Erfolg = False: aus.Status = "FEHLER: " & message
End Sub
