Attribute VB_Name = "modMain"
'==============================================================================
' modMain - buttons of the Pricer sheet
'
' Pricing of every structure is done by live worksheet formulas (Schedule sheet,
' Pricer table, Scenarios grid). VBA only does what a formula cannot:
'   SolveZeroCost     solves the zero-cost strikes (collar and three-way sold option,
'                     participating swap, double-up swap) and writes them into the sheet
'   RunMonteCarlo     outcome statistics of every structure + double-up anatomy
'   PriceAll          both of the above
'   SetDefaultStrike  protection strike = budget if out-of-the-money, else futures +/- 5%;
'                     three-way cap 20% further away
'==============================================================================
Option Explicit

Private dMinY As Double, dMaxY As Double

Private Sub AppFast(ByVal fast As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = Not fast
    If fast Then
        Application.Calculation = xlCalculationManual
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.Calculate
    End If
    On Error GoTo 0
End Sub

Private Sub SetStatus(ByVal msg As String)
    NamedCell("Status").Value = msg
    On Error Resume Next
    Application.StatusBar = msg
    On Error GoTo 0
End Sub

Private Function Stamp() As String
    Stamp = Format(Now, "yyyy-mm-dd hh:mm")
End Function

Public Sub SetDefaultStrike()
    LoadMarket
    SetupDeal
    If gPhi * (gBudget - gFbar) > 0.002 * gFbar Then
        NamedCell("Protect").Value = Round(gBudget, 2)
    Else
        NamedCell("Protect").Value = Round(gFbar * (1# + 0.05 * gPhi), 2)
    End If
    NamedCell("Cap3").Value = Round(NamedValue("Protect") * (1# + 0.2 * gPhi), 2)
End Sub

Public Sub PriceAll()
    Dim t0 As Double, m1 As String, m2 As String
    t0 = Timer
    m1 = DoSolve()
    m2 = DoMonteCarlo()
    SetStatus "Priced " & Stamp() & " | " & m1 & " | " & m2 & " | " & Format(Timer - t0, "0.0") & " s"
End Sub

Public Sub SolveZeroCost()
    Dim t0 As Double, m As String
    t0 = Timer
    m = DoSolve()
    SetStatus "Strikes solved " & Stamp() & " | " & m & " | " & Format(Timer - t0, "0.0") & _
              " s | Monte Carlo statistics not refreshed"
End Sub

Public Sub RunMonteCarlo()
    Dim t0 As Double, m As String
    t0 = Timer
    m = DoMonteCarlo()
    SetStatus "Monte Carlo " & Stamp() & " | " & m & " | " & Format(Timer - t0, "0.0") & " s"
End Sub

' Zero-cost strikes by bisection on the Black-76 strip
Private Function DoSolve() As String
    Dim ok As Boolean, msg As String, k As Double
    On Error GoTo Fail
    AppFast True
    SetStatus "Solving zero-cost strikes..."
    LoadMarket
    SetupDeal
    k = SolveGiveup(1, ok)
    If ok Then NamedCell("Giveup").Value = Round(k, 4) Else msg = msg & "no zero-cost collar; "
    k = SolveGiveup(2, ok)
    If ok Then NamedCell("Giveup3").Value = Round(k, 4) Else msg = msg & "no zero-cost three-way; "
    k = SolveSwapLike(3, ok)
    If ok Then NamedCell("PSStrike").Value = Round(k, 4) Else msg = msg & "no zero-cost participating swap; "
    k = SolveSwapLike(4, ok)
    If ok Then NamedCell("DUStrike").Value = Round(k, 4) Else msg = msg & "no zero-cost double-up; "
    AppFast False
    DoSolve = IIf(Len(msg) = 0, "zero-cost strikes OK", msg)
    Exit Function
Fail:
    AppFast False
    DoSolve = "Error in solver: " & Err.Description
End Function

' Achieved average price (premium included) of strategy `code` on path p of S
Private Function PathPrice(S() As Double, ByVal p As Long, ByVal code As Long, ByVal premBbl As Double, _
                           nDoubled As Long) As Double
    Dim i As Long, v As Double
    nDoubled = 0
    For i = 1 To gNFix
        v = v + S(p, i) - gPhi * LegGain(code, S(p, i), i)
        If code = 7 Then
            If gPhi * (S(p, i) - gDUK) < 0 Then nDoubled = nDoubled + 1
        End If
    Next i
    PathPrice = v / gNFix + gPhi * premBbl
End Function

' Outcome statistics of every structure on simulated paths
Private Function DoMonteCarlo() As String
    Dim nP As Long, seed As Long, onCurve As Boolean
    Dim S() As Double
    Dim code As Long, p As Long, j As Long
    Dim px() As Double, nd As Long, nOver As Long, nWorse As Long, sumX As Double
    Dim premBbl As Double, tAvg As Double
    Dim res(1 To 7, 1 To 5) As Variant
    Dim dist() As Variant, q As Double
    Dim qBad As Double, qGood As Double
    Dim nAny As Long, nAll As Long, sumD As Double

    On Error GoTo Fail
    AppFast True
    SetStatus "Monte Carlo: outcome statistics..."
    LoadMarket
    SetupDeal
    nP = CLng(NamedValue("NPaths"))
    If nP < 1000 Then nP = 1000
    seed = CLng(NamedValue("Seed"))
    onCurve = (InStr(1, CStr(NamedValue("DriftView")), "curve", vbTextCompare) > 0)
    gPremium = StripOption(gProtect, ProtectIsCall(), gQ)
    tAvg = (gT(1) + gT(gNFix)) / 2#
    ' premium carried to the average fixing date, spread over all the barrels
    premBbl = gPremium / DF(tAvg) / (gQ * gNFix)

    SimulatePaths S, nP, seed, onCurve
    nP = UBound(S, 1)
    If gPhi > 0 Then qBad = 0.95: qGood = 0.05 Else qBad = 0.05: qGood = 0.95
    ReDim dist(1 To 99, 1 To N_STRAT + 1)
    For j = 1 To 99
        dist(j, 1) = j / 100#
    Next j
    dMinY = 1E+30: dMaxY = -1E+30
    For code = 1 To N_STRAT
        ReDim px(1 To nP)
        nOver = 0: nWorse = 0: sumX = 0
        For p = 1 To nP
            If code = 3 Then
                px(p) = PathPrice(S, p, code, premBbl, nd)
            Else
                px(p) = PathPrice(S, p, code, 0#, nd)
            End If
            sumX = sumX + px(p)
            If nd > 0 Then nOver = nOver + 1
            If code = 7 Then
                If nd > 0 Then nAny = nAny + 1
                If nd = gNFix Then nAll = nAll + 1
                sumD = sumD + nd
            End If
            ' 1e-9 tolerance: a price exactly at the budget (cap level) is not a miss
            If gPhi > 0 Then
                If px(p) > gBudget + 0.000000001 Then nWorse = nWorse + 1
            Else
                If px(p) < gBudget - 0.000000001 Then nWorse = nWorse + 1
            End If
        Next p
        QuickSort px, 1, nP
        res(code, 1) = sumX / nP
        res(code, 2) = PercentileSorted(px, qBad)
        res(code, 3) = PercentileSorted(px, qGood)
        res(code, 4) = nWorse / nP
        res(code, 5) = nOver / nP
        For j = 1 To 99
            ' quantile function, ordered from the client's best to worst outcome
            If gPhi > 0 Then q = j / 100# Else q = 1# - j / 100#
            dist(j, code + 1) = PercentileSorted(px, q)
            If dist(j, code + 1) < dMinY Then dMinY = dist(j, code + 1)
            If dist(j, code + 1) > dMaxY Then dMaxY = dist(j, code + 1)
        Next j
    Next code
    Erase S

    NamedCell("McTop").Resize(7, 5).Value = res
    With NamedCell("DUTop")
        .Offset(0, 0).Value = nAny / nP
        .Offset(1, 0).Value = sumD / nP
        .Offset(2, 0).Value = nAll / nP
        .Offset(3, 0).Value = (gLev - 1#) * gQ * sumD / nP
    End With
    NamedCell("DistTop").Resize(99, N_STRAT + 1).Value = dist
    AppFast False
    RescaleCharts
    DoMonteCarlo = nP & " paths"
    Exit Function
Fail:
    AppFast False
    DoMonteCarlo = "Error in Monte Carlo: " & Err.Description
End Function

' Axis ranges follow the data (Excel would otherwise start the price axis at zero)
Public Sub RescaleCharts()
    Dim rng As Range, cell As Range, lo As Double, hi As Double, xlo As Double, xhi As Double
    On Error Resume Next
    lo = 1E+30: hi = -1E+30
    Set rng = NamedCell("ScenTop").Offset(0, 2).Resize(61, N_STRAT)
    For Each cell In rng.Cells
        If IsNumeric(cell.Value) Then
            If cell.Value < lo Then lo = cell.Value
            If cell.Value > hi Then hi = cell.Value
        End If
    Next cell
    xlo = NamedCell("ScenTop").Offset(0, 1).Value
    xhi = NamedCell("ScenTop").Offset(60, 1).Value
    With ThisWorkbook.Worksheets("Pricer").ChartObjects(1).Chart
        .Axes(1).MinimumScale = Int(xlo / 10#) * 10#
        .Axes(1).MaximumScale = -Int(-xhi / 10#) * 10#
        .Axes(2).MinimumScale = Int(lo / 10#) * 10#
        .Axes(2).MaximumScale = -Int(-hi / 10#) * 10#
    End With
    If dMaxY > dMinY Then
        With ThisWorkbook.Worksheets("Distribution").ChartObjects(1).Chart
            .Axes(2).MinimumScale = Int(dMinY / 5#) * 5#
            .Axes(2).MaximumScale = -Int(-dMaxY / 5#) * 5#
        End With
    End If
    On Error GoTo 0
End Sub
