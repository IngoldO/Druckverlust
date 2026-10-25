Attribute VB_Name = "NetzValidierung"
Option Explicit

Private Sub Nah(ByVal a As Double, ByVal b As Double, ByVal txt As String)
    If Abs(a-b) > 0.00002 + 0.000001 * Abs(b) Then NetzFehler "Referenzabweichung: " & txt
End Sub

Public Sub NetzwerkValidieren()
    Dim k() As NetzKnoten, e() As NetzLeitung, pq(1 To 2) As Double, ph(1 To 2) As Double
    Dim r As NetzResultat, count As Long, i As Long
    On Error GoTo Fehler
    ReDim k(1 To 2): ReDim e(1 To 1)
    ' single
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 2.98#
    e(1).ID = "L1": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "single: " & r.Status
    Nah r.Druck(1), 3#, "single Druck 1"
    Nah r.Druck(2), 2.98#, "single Druck 2"
    Nah r.Q(1), 1.1044661672776628#, "single Q 1"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 3): ReDim e(1 To 2)
    ' series
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "M": k(2).Hoehe = 0#
    k(3).ID = "B": k(3).Hoehe = 0#
    k(3).DruckFest = True: k(3).DruckBarA = 2.98#
    e(1).ID = "L1": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    e(2).ID = "L2": e(2).Von = 2: e(2).Nach = 3
    e(2).D = 0.050000000000000003#: e(2).L = 10#: e(2).Eps = 0#
    e(2).Zeta = 0#: e(2).ZKv = 0#: e(2).ZBlende = 0#
    e(2).Aktiv = True: e(2).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "series: " & r.Status
    Nah r.Druck(1), 3#, "series Druck 1"
    Nah r.Druck(2), 2.9900000000000002#, "series Druck 2"
    Nah r.Druck(3), 2.98#, "series Druck 3"
    Nah r.Q(1), 0.55223308363881918#, "series Q 1"
    Nah r.Q(2), 0.5522330836388436#, "series Q 2"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 2)
    ' parallel_identical
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 2.98#
    e(1).ID = "L0": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    e(2).ID = "L1": e(2).Von = 1: e(2).Nach = 2
    e(2).D = 0.050000000000000003#: e(2).L = 10#: e(2).Eps = 0#
    e(2).Zeta = 0#: e(2).ZKv = 0#: e(2).ZBlende = 0#
    e(2).Aktiv = True: e(2).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "parallel_identical: " & r.Status
    Nah r.Druck(1), 3#, "parallel_identical Druck 1"
    Nah r.Druck(2), 2.98#, "parallel_identical Druck 2"
    Nah r.Q(1), 1.1044661672776628#, "parallel_identical Q 1"
    Nah r.Q(2), 1.1044661672776628#, "parallel_identical Q 2"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 2)
    ' parallel_different
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 2.98#
    e(1).ID = "L0": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    e(2).ID = "L1": e(2).Von = 1: e(2).Nach = 2
    e(2).D = 0.050000000000000003#: e(2).L = 20#: e(2).Eps = 0#
    e(2).Zeta = 0#: e(2).ZKv = 0#: e(2).ZBlende = 0#
    e(2).Aktiv = True: e(2).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "parallel_different: " & r.Status
    Nah r.Druck(1), 3#, "parallel_different Druck 1"
    Nah r.Druck(2), 2.98#, "parallel_different Druck 2"
    Nah r.Q(1), 1.1044661672776628#, "parallel_different Q 1"
    Nah r.Q(2), 0.55223308363883139#, "parallel_different Q 2"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 3)
    ' parallel_three
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 2.98#
    e(1).ID = "L0": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    e(2).ID = "L1": e(2).Von = 1: e(2).Nach = 2
    e(2).D = 0.050000000000000003#: e(2).L = 20#: e(2).Eps = 0#
    e(2).Zeta = 0#: e(2).ZKv = 0#: e(2).ZBlende = 0#
    e(2).Aktiv = True: e(2).Rueckschlag = False
    e(3).ID = "L2": e(3).Von = 1: e(3).Nach = 2
    e(3).D = 0.050000000000000003#: e(3).L = 30#: e(3).Eps = 0#
    e(3).Zeta = 0#: e(3).ZKv = 0#: e(3).ZBlende = 0#
    e(3).Aktiv = True: e(3).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "parallel_three: " & r.Status
    Nah r.Druck(1), 3#, "parallel_three Druck 1"
    Nah r.Druck(2), 2.98#, "parallel_three Druck 2"
    Nah r.Q(1), 1.1044661672776628#, "parallel_three Q 1"
    Nah r.Q(2), 0.55223308363883139#, "parallel_three Q 2"
    Nah r.Q(3), 0.36815538909255424#, "parallel_three Q 3"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 6): ReDim e(1 To 6)
    ' collector
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "V": k(2).Hoehe = 0#
    k(3).ID = "X": k(3).Hoehe = 0#
    k(4).ID = "Y": k(4).Hoehe = 0#
    k(5).ID = "S": k(5).Hoehe = 0#
    k(6).ID = "B": k(6).Hoehe = 0#
    k(6).DruckFest = True: k(6).DruckBarA = 2.98#
    e(1).ID = "in": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    e(2).ID = "x1": e(2).Von = 2: e(2).Nach = 3
    e(2).D = 0.050000000000000003#: e(2).L = 10#: e(2).Eps = 0#
    e(2).Zeta = 0#: e(2).ZKv = 0#: e(2).ZBlende = 0#
    e(2).Aktiv = True: e(2).Rueckschlag = False
    e(3).ID = "y1": e(3).Von = 2: e(3).Nach = 4
    e(3).D = 0.050000000000000003#: e(3).L = 20#: e(3).Eps = 0#
    e(3).Zeta = 0#: e(3).ZKv = 0#: e(3).ZBlende = 0#
    e(3).Aktiv = True: e(3).Rueckschlag = False
    e(4).ID = "x2": e(4).Von = 3: e(4).Nach = 5
    e(4).D = 0.050000000000000003#: e(4).L = 10#: e(4).Eps = 0#
    e(4).Zeta = 0#: e(4).ZKv = 0#: e(4).ZBlende = 0#
    e(4).Aktiv = True: e(4).Rueckschlag = False
    e(5).ID = "y2": e(5).Von = 4: e(5).Nach = 5
    e(5).D = 0.050000000000000003#: e(5).L = 10#: e(5).Eps = 0#
    e(5).Zeta = 0#: e(5).ZKv = 0#: e(5).ZBlende = 0#
    e(5).Aktiv = True: e(5).Rueckschlag = False
    e(6).ID = "out": e(6).Von = 5: e(6).Nach = 6
    e(6).D = 0.050000000000000003#: e(6).L = 10#: e(6).Eps = 0#
    e(6).Zeta = 0#: e(6).ZKv = 0#: e(6).ZBlende = 0#
    e(6).Aktiv = True: e(6).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "collector: " & r.Status
    Nah r.Druck(1), 3#, "collector Druck 1"
    Nah r.Druck(2), 2.9937499999999999#, "collector Druck 2"
    Nah r.Druck(3), 2.9900000000000002#, "collector Druck 3"
    Nah r.Druck(4), 2.98875#, "collector Druck 4"
    Nah r.Druck(5), 2.9862500000000001#, "collector Druck 5"
    Nah r.Druck(6), 2.98#, "collector Druck 6"
    Nah r.Q(1), 0.34514567727427425#, "collector Q 1"
    Nah r.Q(2), 0.20708740636454492#, "collector Q 2"
    Nah r.Q(3), 0.13805827090970479#, "collector Q 3"
    Nah r.Q(4), 0.20708740636456943#, "collector Q 4"
    Nah r.Q(5), 0.13805827090970479#, "collector Q 5"
    Nah r.Q(6), 0.34514567727427425#, "collector Q 6"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 4): ReDim e(1 To 3)
    ' consumers
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "V": k(2).Hoehe = 0#
    k(3).ID = "B": k(3).Hoehe = 0#
    k(3).DruckFest = True: k(3).DruckBarA = 2.98#
    k(4).ID = "C": k(4).Hoehe = 0#
    k(4).FlussVorgegeben = True: k(4).ExternQ = -0.10000000000000001#
    e(1).ID = "in": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    e(2).ID = "one": e(2).Von = 2: e(2).Nach = 3
    e(2).D = 0.050000000000000003#: e(2).L = 10#: e(2).Eps = 0#
    e(2).Zeta = 0#: e(2).ZKv = 0#: e(2).ZBlende = 0#
    e(2).Aktiv = True: e(2).Rueckschlag = False
    e(3).ID = "two": e(3).Von = 2: e(3).Nach = 4
    e(3).D = 0.050000000000000003#: e(3).L = 10#: e(3).Eps = 0#
    e(3).Zeta = 0#: e(3).ZKv = 0#: e(3).ZBlende = 0#
    e(3).Aktiv = True: e(3).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "consumers: " & r.Status
    Nah r.Druck(1), 3#, "consumers Druck 1"
    Nah r.Druck(2), 2.9890945852126327#, "consumers Druck 2"
    Nah r.Druck(3), 2.98#, "consumers Druck 3"
    Nah r.Druck(4), 2.9872837556378982#, "consumers Druck 4"
    Nah r.Q(1), 0.6022330836388341#, "consumers Q 1"
    Nah r.Q(2), 0.50223308363882868#, "consumers Q 2"
    Nah r.Q(3), 0.10000000000000542#, "consumers Q 3"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 4): ReDim e(1 To 5)
    ' loop
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "X": k(2).Hoehe = 0#
    k(3).ID = "Y": k(3).Hoehe = 0#
    k(4).ID = "B": k(4).Hoehe = 0#
    k(4).DruckFest = True: k(4).DruckBarA = 2.98#
    e(1).ID = "ax": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    e(2).ID = "ay": e(2).Von = 1: e(2).Nach = 3
    e(2).D = 0.050000000000000003#: e(2).L = 20#: e(2).Eps = 0#
    e(2).Zeta = 0#: e(2).ZKv = 0#: e(2).ZBlende = 0#
    e(2).Aktiv = True: e(2).Rueckschlag = False
    e(3).ID = "xb": e(3).Von = 2: e(3).Nach = 4
    e(3).D = 0.050000000000000003#: e(3).L = 30#: e(3).Eps = 0#
    e(3).Zeta = 0#: e(3).ZKv = 0#: e(3).ZBlende = 0#
    e(3).Aktiv = True: e(3).Rueckschlag = False
    e(4).ID = "yb": e(4).Von = 3: e(4).Nach = 4
    e(4).D = 0.050000000000000003#: e(4).L = 10#: e(4).Eps = 0#
    e(4).Zeta = 0#: e(4).ZKv = 0#: e(4).ZBlende = 0#
    e(4).Aktiv = True: e(4).Rueckschlag = False
    e(5).ID = "xy": e(5).Von = 2: e(5).Nach = 3
    e(5).D = 0.050000000000000003#: e(5).L = 10#: e(5).Eps = 0#
    e(5).Zeta = 0#: e(5).ZKv = 0#: e(5).ZBlende = 0#
    e(5).Aktiv = True: e(5).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "loop: " & r.Status
    Nah r.Druck(1), 3#, "loop Druck 1"
    Nah r.Druck(2), 2.9924137931034482#, "loop Druck 2"
    Nah r.Druck(3), 2.9889655172413794#, "loop Druck 3"
    Nah r.Druck(4), 2.98#, "loop Druck 4"
    Nah r.Q(1), 0.41893544276049449#, "loop Q 1"
    Nah r.Q(2), 0.30468032200762901#, "loop Q 2"
    Nah r.Q(3), 0.22851024150572277#, "loop Q 3"
    Nah r.Q(4), 0.49510552326240481#, "loop Q 4"
    Nah r.Q(5), 0.19042520125476353#, "loop Q 5"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    ' orifice
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 2.98#
    e(1).ID = "L": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 29.456901619036948#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "orifice: " & r.Status
    Nah r.Druck(1), 3#, "orifice Druck 1"
    Nah r.Druck(2), 2.98#, "orifice Druck 2"
    Nah r.Q(1), 0.95576432735066641#, "orifice Q 1"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    ' kv
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 2.98#
    e(1).ID = "L": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 99.929744561029779#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "kv: " & r.Status
    Nah r.Druck(1), 3#, "kv Druck 1"
    Nah r.Druck(2), 2.98#, "kv Druck 2"
    Nah r.Q(1), 0.77380405994932788#, "kv Q 1"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    ' zero
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 3#
    e(1).ID = "L": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "zero: " & r.Status
    Nah r.Druck(1), 3#, "zero Druck 1"
    Nah r.Druck(2), 3#, "zero Druck 2"
    Nah r.Q(1), 0#, "zero Q 1"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    ' reverse
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 2.98#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 3#
    e(1).ID = "L": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "reverse: " & r.Status
    Nah r.Druck(1), 2.98#, "reverse Druck 1"
    Nah r.Druck(2), 3#, "reverse Druck 2"
    Nah r.Q(1), -1.1044661672776628#, "reverse Q 1"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    ' elevation
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 10#
    k(2).DruckFest = True: k(2).DruckBarA = 2.0193349999999999#
    e(1).ID = "L": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "elevation: " & r.Status
    Nah r.Druck(1), 3#, "elevation Druck 1"
    Nah r.Druck(2), 2.0193349999999999#, "elevation Druck 2"
    Nah r.Q(1), 0#, "elevation Q 1"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 20): ReDim e(1 To 30)
    ' mesh20x30
    k(1).ID = "0": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "1": k(2).Hoehe = 0#
    k(3).ID = "2": k(3).Hoehe = 0#
    k(4).ID = "3": k(4).Hoehe = 0#
    k(5).ID = "4": k(5).Hoehe = 0#
    k(6).ID = "5": k(6).Hoehe = 0#
    k(7).ID = "6": k(7).Hoehe = 0#
    k(8).ID = "7": k(8).Hoehe = 0#
    k(9).ID = "8": k(9).Hoehe = 0#
    k(10).ID = "9": k(10).Hoehe = 0#
    k(11).ID = "10": k(11).Hoehe = 0#
    k(12).ID = "11": k(12).Hoehe = 0#
    k(13).ID = "12": k(13).Hoehe = 0#
    k(14).ID = "13": k(14).Hoehe = 0#
    k(15).ID = "14": k(15).Hoehe = 0#
    k(16).ID = "15": k(16).Hoehe = 0#
    k(17).ID = "16": k(17).Hoehe = 0#
    k(18).ID = "17": k(18).Hoehe = 0#
    k(19).ID = "18": k(19).Hoehe = 0#
    k(20).ID = "19": k(20).Hoehe = 0#
    k(20).DruckFest = True: k(20).DruckBarA = 2.7999999999999998#
    e(1).ID = "chain0": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    e(2).ID = "chain1": e(2).Von = 2: e(2).Nach = 3
    e(2).D = 0.050000000000000003#: e(2).L = 10#: e(2).Eps = 0#
    e(2).Zeta = 0#: e(2).ZKv = 0#: e(2).ZBlende = 0#
    e(2).Aktiv = True: e(2).Rueckschlag = False
    e(3).ID = "chain2": e(3).Von = 3: e(3).Nach = 4
    e(3).D = 0.050000000000000003#: e(3).L = 10#: e(3).Eps = 0#
    e(3).Zeta = 0#: e(3).ZKv = 0#: e(3).ZBlende = 0#
    e(3).Aktiv = True: e(3).Rueckschlag = False
    e(4).ID = "chain3": e(4).Von = 4: e(4).Nach = 5
    e(4).D = 0.050000000000000003#: e(4).L = 10#: e(4).Eps = 0#
    e(4).Zeta = 0#: e(4).ZKv = 0#: e(4).ZBlende = 0#
    e(4).Aktiv = True: e(4).Rueckschlag = False
    e(5).ID = "chain4": e(5).Von = 5: e(5).Nach = 6
    e(5).D = 0.050000000000000003#: e(5).L = 10#: e(5).Eps = 0#
    e(5).Zeta = 0#: e(5).ZKv = 0#: e(5).ZBlende = 0#
    e(5).Aktiv = True: e(5).Rueckschlag = False
    e(6).ID = "chain5": e(6).Von = 6: e(6).Nach = 7
    e(6).D = 0.050000000000000003#: e(6).L = 10#: e(6).Eps = 0#
    e(6).Zeta = 0#: e(6).ZKv = 0#: e(6).ZBlende = 0#
    e(6).Aktiv = True: e(6).Rueckschlag = False
    e(7).ID = "chain6": e(7).Von = 7: e(7).Nach = 8
    e(7).D = 0.050000000000000003#: e(7).L = 10#: e(7).Eps = 0#
    e(7).Zeta = 0#: e(7).ZKv = 0#: e(7).ZBlende = 0#
    e(7).Aktiv = True: e(7).Rueckschlag = False
    e(8).ID = "chain7": e(8).Von = 8: e(8).Nach = 9
    e(8).D = 0.050000000000000003#: e(8).L = 10#: e(8).Eps = 0#
    e(8).Zeta = 0#: e(8).ZKv = 0#: e(8).ZBlende = 0#
    e(8).Aktiv = True: e(8).Rueckschlag = False
    e(9).ID = "chain8": e(9).Von = 9: e(9).Nach = 10
    e(9).D = 0.050000000000000003#: e(9).L = 10#: e(9).Eps = 0#
    e(9).Zeta = 0#: e(9).ZKv = 0#: e(9).ZBlende = 0#
    e(9).Aktiv = True: e(9).Rueckschlag = False
    e(10).ID = "chain9": e(10).Von = 10: e(10).Nach = 11
    e(10).D = 0.050000000000000003#: e(10).L = 10#: e(10).Eps = 0#
    e(10).Zeta = 0#: e(10).ZKv = 0#: e(10).ZBlende = 0#
    e(10).Aktiv = True: e(10).Rueckschlag = False
    e(11).ID = "chain10": e(11).Von = 11: e(11).Nach = 12
    e(11).D = 0.050000000000000003#: e(11).L = 10#: e(11).Eps = 0#
    e(11).Zeta = 0#: e(11).ZKv = 0#: e(11).ZBlende = 0#
    e(11).Aktiv = True: e(11).Rueckschlag = False
    e(12).ID = "chain11": e(12).Von = 12: e(12).Nach = 13
    e(12).D = 0.050000000000000003#: e(12).L = 10#: e(12).Eps = 0#
    e(12).Zeta = 0#: e(12).ZKv = 0#: e(12).ZBlende = 0#
    e(12).Aktiv = True: e(12).Rueckschlag = False
    e(13).ID = "chain12": e(13).Von = 13: e(13).Nach = 14
    e(13).D = 0.050000000000000003#: e(13).L = 10#: e(13).Eps = 0#
    e(13).Zeta = 0#: e(13).ZKv = 0#: e(13).ZBlende = 0#
    e(13).Aktiv = True: e(13).Rueckschlag = False
    e(14).ID = "chain13": e(14).Von = 14: e(14).Nach = 15
    e(14).D = 0.050000000000000003#: e(14).L = 10#: e(14).Eps = 0#
    e(14).Zeta = 0#: e(14).ZKv = 0#: e(14).ZBlende = 0#
    e(14).Aktiv = True: e(14).Rueckschlag = False
    e(15).ID = "chain14": e(15).Von = 15: e(15).Nach = 16
    e(15).D = 0.050000000000000003#: e(15).L = 10#: e(15).Eps = 0#
    e(15).Zeta = 0#: e(15).ZKv = 0#: e(15).ZBlende = 0#
    e(15).Aktiv = True: e(15).Rueckschlag = False
    e(16).ID = "chain15": e(16).Von = 16: e(16).Nach = 17
    e(16).D = 0.050000000000000003#: e(16).L = 10#: e(16).Eps = 0#
    e(16).Zeta = 0#: e(16).ZKv = 0#: e(16).ZBlende = 0#
    e(16).Aktiv = True: e(16).Rueckschlag = False
    e(17).ID = "chain16": e(17).Von = 17: e(17).Nach = 18
    e(17).D = 0.050000000000000003#: e(17).L = 10#: e(17).Eps = 0#
    e(17).Zeta = 0#: e(17).ZKv = 0#: e(17).ZBlende = 0#
    e(17).Aktiv = True: e(17).Rueckschlag = False
    e(18).ID = "chain17": e(18).Von = 18: e(18).Nach = 19
    e(18).D = 0.050000000000000003#: e(18).L = 10#: e(18).Eps = 0#
    e(18).Zeta = 0#: e(18).ZKv = 0#: e(18).ZBlende = 0#
    e(18).Aktiv = True: e(18).Rueckschlag = False
    e(19).ID = "chain18": e(19).Von = 19: e(19).Nach = 20
    e(19).D = 0.050000000000000003#: e(19).L = 10#: e(19).Eps = 0#
    e(19).Zeta = 0#: e(19).ZKv = 0#: e(19).ZBlende = 0#
    e(19).Aktiv = True: e(19).Rueckschlag = False
    e(20).ID = "cross0": e(20).Von = 1: e(20).Nach = 3
    e(20).D = 0.050000000000000003#: e(20).L = 15#: e(20).Eps = 0#
    e(20).Zeta = 0#: e(20).ZKv = 0#: e(20).ZBlende = 0#
    e(20).Aktiv = True: e(20).Rueckschlag = False
    e(21).ID = "cross1": e(21).Von = 2: e(21).Nach = 4
    e(21).D = 0.050000000000000003#: e(21).L = 16#: e(21).Eps = 0#
    e(21).Zeta = 0#: e(21).ZKv = 0#: e(21).ZBlende = 0#
    e(21).Aktiv = True: e(21).Rueckschlag = False
    e(22).ID = "cross2": e(22).Von = 3: e(22).Nach = 5
    e(22).D = 0.050000000000000003#: e(22).L = 17#: e(22).Eps = 0#
    e(22).Zeta = 0#: e(22).ZKv = 0#: e(22).ZBlende = 0#
    e(22).Aktiv = True: e(22).Rueckschlag = False
    e(23).ID = "cross3": e(23).Von = 4: e(23).Nach = 6
    e(23).D = 0.050000000000000003#: e(23).L = 18#: e(23).Eps = 0#
    e(23).Zeta = 0#: e(23).ZKv = 0#: e(23).ZBlende = 0#
    e(23).Aktiv = True: e(23).Rueckschlag = False
    e(24).ID = "cross4": e(24).Von = 5: e(24).Nach = 7
    e(24).D = 0.050000000000000003#: e(24).L = 19#: e(24).Eps = 0#
    e(24).Zeta = 0#: e(24).ZKv = 0#: e(24).ZBlende = 0#
    e(24).Aktiv = True: e(24).Rueckschlag = False
    e(25).ID = "cross5": e(25).Von = 6: e(25).Nach = 8
    e(25).D = 0.050000000000000003#: e(25).L = 20#: e(25).Eps = 0#
    e(25).Zeta = 0#: e(25).ZKv = 0#: e(25).ZBlende = 0#
    e(25).Aktiv = True: e(25).Rueckschlag = False
    e(26).ID = "cross6": e(26).Von = 7: e(26).Nach = 9
    e(26).D = 0.050000000000000003#: e(26).L = 21#: e(26).Eps = 0#
    e(26).Zeta = 0#: e(26).ZKv = 0#: e(26).ZBlende = 0#
    e(26).Aktiv = True: e(26).Rueckschlag = False
    e(27).ID = "cross7": e(27).Von = 8: e(27).Nach = 10
    e(27).D = 0.050000000000000003#: e(27).L = 22#: e(27).Eps = 0#
    e(27).Zeta = 0#: e(27).ZKv = 0#: e(27).ZBlende = 0#
    e(27).Aktiv = True: e(27).Rueckschlag = False
    e(28).ID = "cross8": e(28).Von = 9: e(28).Nach = 11
    e(28).D = 0.050000000000000003#: e(28).L = 23#: e(28).Eps = 0#
    e(28).Zeta = 0#: e(28).ZKv = 0#: e(28).ZBlende = 0#
    e(28).Aktiv = True: e(28).Rueckschlag = False
    e(29).ID = "cross9": e(29).Von = 10: e(29).Nach = 12
    e(29).D = 0.050000000000000003#: e(29).L = 24#: e(29).Eps = 0#
    e(29).Zeta = 0#: e(29).ZKv = 0#: e(29).ZBlende = 0#
    e(29).Aktiv = True: e(29).Rueckschlag = False
    e(30).ID = "cross10": e(30).Von = 11: e(30).Nach = 13
    e(30).D = 0.050000000000000003#: e(30).L = 25#: e(30).Eps = 0#
    e(30).Zeta = 0#: e(30).ZKv = 0#: e(30).ZBlende = 0#
    e(30).Aktiv = True: e(30).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "mesh20x30: " & r.Status
    Nah r.Druck(1), 3#, "mesh20x30 Druck 1"
    Nah r.Druck(2), 2.9908302050977551#, "mesh20x30 Druck 2"
    Nah r.Druck(3), 2.9873336228305507#, "mesh20x30 Druck 3"
    Nah r.Druck(4), 2.9817530648816897#, "mesh20x30 Druck 4"
    Nah r.Druck(5), 2.9765211540306589#, "mesh20x30 Druck 5"
    Nah r.Druck(6), 2.9709137173625217#, "mesh20x30 Druck 6"
    Nah r.Druck(7), 2.9651501291303388#, "mesh20x30 Druck 7"
    Nah r.Druck(8), 2.9591823010248719#, "mesh20x30 Druck 8"
    Nah r.Druck(9), 2.9530110579744089#, "mesh20x30 Druck 9"
    Nah r.Druck(10), 2.9467252559324479#, "mesh20x30 Druck 10"
    Nah r.Druck(11), 2.9399793704840738#, "mesh20x30 Druck 11"
    Nah r.Druck(12), 2.9342399523706502#, "mesh20x30 Druck 12"
    Nah r.Druck(13), 2.923298324439811#, "mesh20x30 Druck 13"
    Nah r.Druck(14), 2.9056842780912664#, "mesh20x30 Druck 14"
    Nah r.Druck(15), 2.8880702317427218#, "mesh20x30 Druck 15"
    Nah r.Druck(16), 2.8704561853941777#, "mesh20x30 Druck 16"
    Nah r.Druck(17), 2.8528421390456331#, "mesh20x30 Druck 17"
    Nah r.Druck(18), 2.8352280926970885#, "mesh20x30 Druck 18"
    Nah r.Druck(19), 2.8176140463485444#, "mesh20x30 Druck 19"
    Nah r.Druck(20), 2.7999999999999998#, "mesh20x30 Druck 20"
    Nah r.Q(1), 0.50638641152023323#, "mesh20x30 Q 1"
    Nah r.Q(2), 0.19309284076151381#, "mesh20x30 Q 2"
    Nah r.Q(3), 0.30817687245247166#, "mesh20x30 Q 3"
    Nah r.Q(4), 0.28892342625882028#, "mesh20x30 Q 4"
    Nah r.Q(5), 0.30966120425548149#, "mesh20x30 Q 5"
    Nah r.Q(6), 0.31828441022828552#, "mesh20x30 Q 6"
    Nah r.Q(7), 0.32956321173084802#, "mesh20x30 Q 7"
    Nah r.Q(8), 0.34079645796418967#, "mesh20x30 Q 8"
    Nah r.Q(9), 0.34712278447753508#, "mesh20x30 Q 9"
    Nah r.Q(10), 0.37253011230299898#, "mesh20x30 Q 10"
    Nah r.Q(11), 0.31694965630684763#, "mesh20x30 Q 11"
    Nah r.Q(12), 0.60423289322760887#, "mesh20x30 Q 12"
    Nah r.Q(13), 0.97270591304140619#, "mesh20x30 Q 13"
    Nah r.Q(14), 0.97270591304140619#, "mesh20x30 Q 14"
    Nah r.Q(15), 0.97270591304138165#, "mesh20x30 Q 15"
    Nah r.Q(16), 0.97270591304140619#, "mesh20x30 Q 16"
    Nah r.Q(17), 0.97270591304140619#, "mesh20x30 Q 17"
    Nah r.Q(18), 0.97270591304138165#, "mesh20x30 Q 18"
    Nah r.Q(19), 0.97270591304140619#, "mesh20x30 Q 19"
    Nah r.Q(20), 0.46631950152116469#, "mesh20x30 Q 20"
    Nah r.Q(21), 0.31329357075874092#, "mesh20x30 Q 21"
    Nah r.Q(22), 0.35123546983017168#, "mesh20x30 Q 22"
    Nah r.Q(23), 0.33254701695238986#, "mesh20x30 Q 23"
    Nah r.Q(24), 0.33049769183356154#, "mesh20x30 Q 24"
    Nah r.Q(25), 0.3239238109795668#, "mesh20x30 Q 25"
    Nah r.Q(26), 0.31921889033097034#, "mesh20x30 Q 26"
    Nah r.Q(27), 0.31269056474623852#, "mesh20x30 Q 27"
    Nah r.Q(28), 0.31289256381762348#, "mesh20x30 Q 28"
    Nah r.Q(29), 0.2872832369207694#, "mesh20x30 Q 29"
    Nah r.Q(30), 0.3684730198137825#, "mesh20x30 Q 30"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    ' check_closed
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 2.98#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 3#
    e(1).ID = "L": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 2#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = True
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "check_closed: " & r.Status
    Nah r.Druck(1), 2.98#, "check_closed Druck 1"
    Nah r.Druck(2), 3#, "check_closed Druck 2"
    Nah r.Q(1), 0#, "check_closed Q 1"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    ' check_open
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 2.98#
    e(1).ID = "L": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 2#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = True
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "check_open: " & r.Status
    Nah r.Druck(1), 3#, "check_open Druck 1"
    Nah r.Druck(2), 2.98#, "check_open Druck 2"
    Nah r.Q(1), 1.0913033569933615#, "check_open Q 1"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 2)
    ' inactive
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 2.98#
    e(1).ID = "on": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 10#: e(1).Eps = 0#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    e(2).ID = "off": e(2).Von = 1: e(2).Nach = 2
    e(2).D = 0.050000000000000003#: e(2).L = 10#: e(2).Eps = 0#
    e(2).Zeta = 0#: e(2).ZKv = 0#: e(2).ZBlende = 0#
    e(2).Aktiv = False: e(2).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.10000000000000001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "inactive: " & r.Status
    Nah r.Druck(1), 3#, "inactive Druck 1"
    Nah r.Druck(2), 2.98#, "inactive Druck 2"
    Nah r.Q(1), 1.1044661672776628#, "inactive Q 1"
    Nah r.Q(2), 0#, "inactive Q 2"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    ' turbulent
    k(1).ID = "A": k(1).Hoehe = 0#
    k(1).DruckFest = True: k(1).DruckBarA = 3#
    k(2).ID = "B": k(2).Hoehe = 0#
    k(2).DruckFest = True: k(2).DruckBarA = 2#
    e(1).ID = "L": e(1).Von = 1: e(1).Nach = 2
    e(1).D = 0.050000000000000003#: e(1).L = 100#: e(1).Eps = 4.5000000000000003E-05#
    e(1).Zeta = 0#: e(1).ZKv = 0#: e(1).ZBlende = 0#
    e(1).Aktiv = True: e(1).Rueckschlag = False
    NetzLoesen k, e, 1000, 0.001#, pq, ph, r
    If Not r.Erfolg Then NetzFehler "turbulent: " & r.Status
    Nah r.Druck(1), 3#, "turbulent Druck 1"
    Nah r.Druck(2), 2#, "turbulent Druck 2"
    Nah r.Q(1), 15.181077520413433#, "turbulent Q 1"
    If r.MaxBilanz > 0.000001 Or r.MaxEnergie > 0.00000001 Then NetzFehler "Residuen ueberschritten."
    count = count + 1
    ReDim k(1 To 2): ReDim e(1 To 1)
    k(1).ID = "A": k(2).ID = "B": k(1).DruckFest = True: k(2).DruckFest = True
    k(1).DruckBarA = 2: k(2).DruckBarA = 2
    e(1).ID = "P": e(1).Von = 1: e(1).Nach = 2: e(1).D = 0.05: e(1).Zeta = 2
    e(1).Aktiv = True: e(1).Pumpe = True
    pq(1) = 0: pq(2) = 40: ph(1) = 4 * 100000# / (1000 * G): ph(2) = 0
    NetzLoesen k, e, 1000, 0.001, pq, ph, r
    If Not r.Erfolg Then NetzFehler "Pumpenkopplung: " & r.Status
    Nah r.Q(1), 37.226436301264144#, "Pumpenbetriebspunkt analytisch"
    count = count + 1
    e(1).Pumpe = False: e(1).L = 10: e(1).Zeta = 0
    k(1).DruckFest = False: k(2).DruckFest = False
    NetzLoesen k, e, 1000, 0.1, pq, ph, r
    If r.Erfolg Or InStr(r.Status, "Druckrandbedingung") = 0 Then NetzFehler "Fehlende Druckvorgabe nicht erkannt."
    count = count + 1
    k(1).DruckFest = True: k(1).FlussVorgegeben = True
    NetzLoesen k, e, 1000, 0.1, pq, ph, r
    If r.Erfolg Or InStr(r.Status, "Ueberbestimmt") = 0 Then NetzFehler "Ueberbestimmung nicht erkannt."
    count = count + 1
    k(1).FlussVorgegeben = False: ReDim Preserve k(1 To 3): k(3).ID = "isoliert"
    NetzLoesen k, e, 1000, 0.1, pq, ph, r
    If r.Erfolg Or InStr(r.Status, "Teilnetz") = 0 Then NetzFehler "Insel nicht erkannt."
    count = count + 1
    ReDim Preserve k(1 To 2): k(2).DruckFest = True: k(1).DruckBarA = 3: k(2).DruckBarA = 2.98
    NetzLoesen k, e, 1000, 0.1, pq, ph, r, 1, 0.00000000000001, 0.00000000000001
    If r.Erfolg Then NetzFehler "Iterationslimit wurde ignoriert."
    count = count + 1
    e(1).L = 0
    NetzLoesen k, e, 1000, 0.1, pq, ph, r
    If r.Erfolg Then NetzFehler "Ideale widerstandslose Leitung akzeptiert."
    count = count + 1
    NetzArmaturAddieren e(1), "Zeta", "Frei definierbarer Einzelwiderstand", 1, 1000, 0
    Nah e(1).Zeta, 0, "Expliziter Nullwert bleibt gueltig"
    count = count + 1
    MsgBox count & " Netzwerk-Testfaelle bestanden.", vbInformation
    Exit Sub
Fehler:
    MsgBox "NETZVALIDIERUNG FEHLGESCHLAGEN: " & Err.Description, vbCritical
    Err.Raise vbObjectError + 250, "NetzValidierung", "Validierung fehlgeschlagen"
End Sub
