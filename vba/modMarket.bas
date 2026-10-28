Attribute VB_Name = "modMarket"
'==============================================================================
' modMarket - Brent market: futures curve, USD discount factors, ATM vol term
'             structure and smile, Black-76 pricing. Also worksheet functions (UDFs).
'
' Market inputs live on the "Market" sheet (named range MktTable):
'   years | futures minus front month (USD/bbl) | USD zero rate % | ATM vol % | 25d RR % | 25d BF %
' The front-month price and the vol shift come from the Pricer sheet (names Front, VolShift),
' so moving the front month shifts the whole curve.
' The worksheet functions (CM_FUTURE, CM_DF, CM_VOL...) feed the live formulas of the
' Schedule and Market sheets.
'==============================================================================
Option Explicit

Public mN As Long
Public mT() As Double, mSpr() As Double, mRd() As Double
Public mAtm() As Double, mRR() As Double, mBF() As Double
Public mFront As Double

Public Function NamedValue(ByVal nm As String) As Variant
    NamedValue = ThisWorkbook.Names(nm).RefersToRange.Value
End Function

Public Function NamedCell(ByVal nm As String) As Range
    Set NamedCell = ThisWorkbook.Names(nm).RefersToRange
End Function

' Reads the market table into module arrays
Public Sub LoadMarket()
    Dim tbl As Range, r As Long, n As Long, volShiftPts As Double
    Set tbl = NamedCell("MktTable")
    n = 0
    For r = 1 To tbl.Rows.Count
        If Len(CStr(tbl.Cells(r, 1).Value)) > 0 Then n = n + 1
    Next r
    mN = n
    ReDim mT(1 To n): ReDim mSpr(1 To n): ReDim mRd(1 To n)
    ReDim mAtm(1 To n): ReDim mRR(1 To n): ReDim mBF(1 To n)
    volShiftPts = CDbl(NamedValue("VolShift"))
    n = 0
    For r = 1 To tbl.Rows.Count
        If Len(CStr(tbl.Cells(r, 1).Value)) > 0 Then
            n = n + 1
            mT(n) = CDbl(tbl.Cells(r, 1).Value)
            mSpr(n) = CDbl(tbl.Cells(r, 2).Value)
            mRd(n) = CDbl(tbl.Cells(r, 3).Value) / 100#
            mAtm(n) = (CDbl(tbl.Cells(r, 4).Value) + volShiftPts) / 100#
            mRR(n) = CDbl(tbl.Cells(r, 5).Value) / 100#
            mBF(n) = CDbl(tbl.Cells(r, 6).Value) / 100#
        End If
    Next r
    mFront = CDbl(NamedValue("Front"))
End Sub

' Futures price for delivery month T = front month + curve shape
Public Function Fut(ByVal T As Double) As Double
    Fut = mFront + Interp(T, mT, mSpr)
End Function

' USD discount factor (continuously compounded zero rate)
Public Function DF(ByVal T As Double) As Double
    DF = Exp(-Interp(T, mT, mRd) * T)
End Function

' ATM vol, linear interpolation in total variance, flat vol outside the pillars
Public Function AtmVol(ByVal T As Double) As Double
    Dim w() As Double, i As Long
    If T < 0.000001 Then T = 0.000001
    If T <= mT(1) Then AtmVol = mAtm(1): Exit Function
    If T >= mT(mN) Then AtmVol = mAtm(mN): Exit Function
    ReDim w(1 To mN)
    For i = 1 To mN
        w(i) = mAtm(i) * mAtm(i) * mT(i)
    Next i
    AtmVol = Sqr(Interp(T, mT, w) / T)
End Function

' Smile: quadratic in call delta x through 3 market points
'   x=0.25 (OTM call) -> ATM + BF + RR/2 ; x=0.5 -> ATM ; x=0.75 (OTM put) -> ATM + BF - RR/2
' Fixed-point iteration because the delta depends on the vol.
Public Function SmileVol(ByVal K As Double, ByVal T As Double) As Double
    Dim atm As Double, rr As Double, bf As Double, F As Double
    Dim b As Double, c As Double, sig As Double, nw As Double, d1 As Double, x As Double
    Dim it As Long
    atm = AtmVol(T)
    rr = Interp(T, mT, mRR)
    bf = Interp(T, mT, mBF)
    F = Fut(T)
    b = -2# * rr: c = 16# * bf
    sig = atm
    For it = 1 To 60
        d1 = (Log(F / K) + 0.5 * sig * sig * T) / (sig * Sqr(T))
        x = NormCDF(d1)
        If x < 0.02 Then x = 0.02
        If x > 0.98 Then x = 0.98
        nw = atm + b * (x - 0.5) + c * (x - 0.5) ^ 2
        If nw < 0.01 Then nw = 0.01
        If Abs(nw - sig) < 0.0000000001 Then sig = nw: Exit For
        sig = 0.5 * sig + 0.5 * nw
    Next it
    SmileVol = sig
End Function

' Black-76 price of an option on a futures price (USD per barrel)
Public Function Black76(ByVal F As Double, ByVal K As Double, ByVal T As Double, _
                        ByVal dfac As Double, ByVal vol As Double, ByVal isCall As Boolean) As Double
    Dim d1 As Double, d2 As Double, vs As Double
    If T < 0.000000000001 Then T = 0.000000000001
    vs = vol * Sqr(T)
    d1 = (Log(F / K) + 0.5 * vol * vol * T) / vs
    d2 = d1 - vs
    If isCall Then
        Black76 = dfac * (F * NormCDF(d1) - K * NormCDF(d2))
    Else
        Black76 = dfac * (K * NormCDF(-d2) - F * NormCDF(-d1))
    End If
End Function

' Price with the market curve and the smile
Public Function PriceMkt(ByVal K As Double, ByVal T As Double, ByVal isCall As Boolean) As Double
    PriceMkt = Black76(Fut(T), K, T, DF(T), SmileVol(K, T), isCall)
End Function

'------------------------------------------------------------------ worksheet UDFs
Public Function CM_FUTURE(ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    CM_FUTURE = Fut(T)
End Function

Public Function CM_DF(ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    CM_DF = DF(T)
End Function

Public Function CM_VOL(ByVal K As Double, ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    CM_VOL = SmileVol(K, T)
End Function

Public Function CM_ATMVOL(ByVal T As Double) As Double
    Application.Volatile
    LoadMarket
    CM_ATMVOL = AtmVol(T)
End Function

Public Function CM_OPTION(ByVal K As Double, ByVal T As Double, ByVal callPut As String) As Double
    Application.Volatile
    LoadMarket
    CM_OPTION = PriceMkt(K, T, UCase(Left(callPut, 1)) = "C")
End Function

' Average futures price over the first n monthly fixings
Public Function CM_AVGFUT(ByVal nMonths As Long) As Double
    Dim i As Long, s As Double
    Application.Volatile
    LoadMarket
    For i = 1 To nMonths
        s = s + Fut(i / 12#)
    Next i
    CM_AVGFUT = s / nMonths
End Function
