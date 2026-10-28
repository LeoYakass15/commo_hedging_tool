Attribute VB_Name = "modMonteCarlo"
'==============================================================================
' modMonteCarlo - log-normal Brent paths on the fixing dates, ATM vol term structure
'   "Futures curve": E[S(t)] = F(t)   (the market's view, used for pricing)
'   "Front month unchanged": E[S(t)] = front month today (no backwardation roll-down)
' Antithetic variates, MRG32k3a generator (same numbers in every Excel version).
'==============================================================================
Option Explicit

Public Sub SimulatePaths(S() As Double, ByVal nPaths As Long, ByVal seed As Long, ByVal onCurve As Boolean)
    Dim half As Long, p As Long, i As Long
    Dim tv() As Double, dv() As Double, ctr() As Double, z As Double
    Dim x As Double, xa As Double
    If nPaths Mod 2 = 1 Then nPaths = nPaths + 1
    half = nPaths \ 2
    ReDim S(1 To nPaths, 1 To gNFix)
    ReDim tv(0 To gNFix): ReDim dv(1 To gNFix): ReDim ctr(1 To gNFix)
    tv(0) = 0
    For i = 1 To gNFix
        tv(i) = AtmVol(gT(i)) ^ 2 * gT(i)
        dv(i) = tv(i) - tv(i - 1)
        If dv(i) < 0.000000000001 Then dv(i) = 0.000000000001
        If onCurve Then ctr(i) = gF(i) Else ctr(i) = mFront
    Next i
    RngSeed seed
    For p = 1 To half
        x = 0: xa = 0
        For i = 1 To gNFix
            z = RngNormal() * Sqr(dv(i))
            x = x + z
            xa = xa - z
            S(p, i) = ctr(i) * Exp(x - 0.5 * tv(i))
            S(p + half, i) = ctr(i) * Exp(xa - 0.5 * tv(i))
        Next i
    Next p
End Sub
