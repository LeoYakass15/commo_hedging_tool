Attribute VB_Name = "modStrategies"
'==============================================================================
' modStrategies - hedging structures on a strip of monthly fixings.
'
' The client buys (consumer, phi = +1) or sells (producer, phi = -1) Q barrels every
' month at the market price S. Each structure pays the client a gain g per barrel:
'       achieved price = S - phi * g
' Swap at K: g = phi * (S - K), so the achieved price is K whatever S.
' Mid value to the client = sum_i Q * DF(t_i) * E_Q[ g_i ]  (Black-76).
'
' Strategy codes: 1 Unhedged, 2 Swap, 3 Call (cap), 4 Zero-cost collar,
'                 5 Three-way collar, 6 Participating swap, 7 Double-up swap
'==============================================================================
Option Explicit

Public Const N_STRAT As Long = 7

' Deal state (filled by SetupDeal)
Public gPhi As Double            ' +1 consumer (buys), -1 producer (sells)
Public gQ As Double              ' barrels per fixing
Public gNFix As Long
Public gT() As Double            ' fixing times (years)
Public gF() As Double            ' futures prices
Public gDF() As Double           ' USD discount factors
Public gSwapK() As Double        ' swap price per fixing, sales margin included
Public gBudget As Double
Public gMargin As Double         ' sales margin, USD per barrel
Public gMarginPV As Double       ' margin in USD, present value
Public gProtect As Double, gGiveup As Double
Public gCap3 As Double, gGiveup3 As Double
Public gPSK As Double, gPart As Double
Public gDUK As Double, gLev As Double
Public gPremium As Double        ' option premium, USD
Public gFbar As Double

Public Sub SetupDeal()
    Dim i As Long
    gPhi = IIf(InStr(1, CStr(NamedValue("Side")), "Consumer", vbTextCompare) > 0, 1#, -1#)
    gQ = CDbl(NamedValue("Volume"))
    gNFix = CLng(NamedValue("NMonths"))
    If gNFix < 1 Then gNFix = 1
    If gNFix > 24 Then gNFix = 24
    gBudget = CDbl(NamedValue("Budget"))
    gMargin = CDbl(NamedValue("MarginUSD"))
    gProtect = CDbl(NamedValue("Protect"))
    gCap3 = CDbl(NamedValue("Cap3"))
    gPart = CDbl(NamedValue("Participation"))
    gLev = CDbl(NamedValue("Leverage"))
    gGiveup = CDbl(NamedValue("Giveup"))
    gGiveup3 = CDbl(NamedValue("Giveup3"))
    gPSK = CDbl(NamedValue("PSStrike"))
    gDUK = CDbl(NamedValue("DUStrike"))
    ReDim gT(1 To gNFix): ReDim gF(1 To gNFix): ReDim gDF(1 To gNFix): ReDim gSwapK(1 To gNFix)
    gFbar = 0: gMarginPV = 0
    For i = 1 To gNFix
        gT(i) = i / 12#
        gF(i) = Fut(gT(i))
        gDF(i) = DF(gT(i))
        gFbar = gFbar + gF(i) / gNFix
        ' the margin makes the swap worse for the client
        gSwapK(i) = gF(i) + gPhi * gMargin
        gMarginPV = gMarginPV + gMargin * gQ * gDF(i)
    Next i
End Sub

' Protection is a call for a consumer (prices up), a put for a producer (prices down)
Public Function ProtectIsCall() As Boolean
    ProtectIsCall = (gPhi > 0)
End Function

' USD value of options on qty barrels per fixing, summed over the strip
Public Function StripOption(ByVal K As Double, ByVal isCall As Boolean, ByVal qty As Double) As Double
    Dim i As Long, s As Double
    For i = 1 To gNFix
        s = s + qty * PriceMkt(K, gT(i), isCall)
    Next i
    StripOption = s
End Function

Public Function CollarPV(ByVal kProt As Double, ByVal kGive As Double) As Double
    CollarPV = StripOption(kProt, ProtectIsCall(), gQ) - StripOption(kGive, Not ProtectIsCall(), gQ)
End Function

' Three-way: protection spread kProt / kCap bought, opposite option kGive sold
Public Function ThreeWayPV(ByVal kProt As Double, ByVal kCap As Double, ByVal kGive As Double) As Double
    ThreeWayPV = StripOption(kProt, ProtectIsCall(), gQ) - StripOption(kCap, ProtectIsCall(), gQ) _
               - StripOption(kGive, Not ProtectIsCall(), gQ)
End Function

' Participating swap: protection on Q, opposite option sold on (1 - participation) x Q
Public Function PSPV(ByVal K As Double, ByVal part As Double) As Double
    PSPV = StripOption(K, ProtectIsCall(), gQ) - StripOption(K, Not ProtectIsCall(), (1# - part) * gQ)
End Function

' Double-up swap: protection on Q, opposite option sold on leverage x Q
Public Function DUPV(ByVal K As Double, ByVal lev As Double) As Double
    DUPV = StripOption(K, ProtectIsCall(), gQ) - StripOption(K, Not ProtectIsCall(), lev * gQ)
End Function

' Value to solve to zero. which: 1 collar give-up, 2 three-way give-up, 3 PS strike, 4 DU strike
Private Function EvalStrike(ByVal which As Long, ByVal x As Double) As Double
    Select Case which
        Case 1: EvalStrike = CollarPV(gProtect, x) + gMarginPV
        Case 2: EvalStrike = ThreeWayPV(gProtect, gCap3, x) + gMarginPV
        Case 3: EvalStrike = PSPV(x, gPart) + gMarginPV
        Case 4: EvalStrike = DUPV(x, gLev) + gMarginPV
    End Select
End Function

Private Function Bisect(ByVal which As Long, ByVal lo As Double, ByVal hi As Double, ok As Boolean) As Double
    Dim flo As Double, fhi As Double, mid As Double, fm As Double, it As Long
    flo = EvalStrike(which, lo): fhi = EvalStrike(which, hi)
    ok = (flo * fhi <= 0)
    If Not ok Then Bisect = 0: Exit Function
    For it = 1 To 60
        mid = 0.5 * (lo + hi)
        fm = EvalStrike(which, mid)
        If fm * flo > 0 Then
            lo = mid: flo = fm
        Else
            hi = mid
        End If
        If hi - lo < 0.0000001 Then Exit For
    Next it
    Bisect = 0.5 * (lo + hi)
End Function

' The sold option is on the other side of the protection strike
Public Function SolveGiveup(ByVal which As Long, ok As Boolean) As Double
    Dim lo As Double, hi As Double
    If gPhi > 0 Then
        lo = 0.2 * gFbar
        hi = IIf(gProtect < gFbar, gProtect, gFbar) * 0.99999
    Else
        lo = IIf(gProtect > gFbar, gProtect, gFbar) * 1.00001
        hi = 3# * gFbar
    End If
    SolveGiveup = Bisect(which, lo, hi, ok)
End Function

Public Function SolveSwapLike(ByVal which As Long, ok As Boolean) As Double
    SolveSwapLike = Bisect(which, 0.5 * gFbar, 1.5 * gFbar, ok)
End Function

' Gain per barrel for the client on fixing i at price S
Public Function LegGain(ByVal code As Long, ByVal S As Double, ByVal i As Long) As Double
    Dim g As Double, x As Double, ph As Double
    ph = gPhi
    Select Case code
        Case 1
            g = 0
        Case 2
            g = ph * (S - gSwapK(i))
        Case 3
            g = Max0(ph * (S - gProtect))
        Case 4
            g = Max0(ph * (S - gProtect)) - Max0(ph * (gGiveup - S))
        Case 5
            g = Max0(ph * (S - gProtect)) - Max0(ph * (S - gCap3)) - Max0(ph * (gGiveup3 - S))
        Case 6
            x = ph * (S - gPSK)
            If x > 0 Then g = x Else g = (1# - gPart) * x
        Case 7
            x = ph * (S - gDUK)
            If x > 0 Then g = x Else g = gLev * x
    End Select
    LegGain = g
End Function

Public Function Max0(ByVal x As Double) As Double
    If x > 0 Then Max0 = x Else Max0 = 0
End Function
