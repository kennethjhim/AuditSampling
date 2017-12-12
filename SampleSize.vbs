'for Attribute Sampling
Function HypgeomSample(risk As Double, pE As Double, pT As Double, P As Double)
    Dim Lt As Double, n As Double, k As Double
    Lt = Application.WorksheetFunction.RoundUp(pT * P, 0)
    If P <= 0 Or risk <= 0 Or risk >= 1 Or pE < 0 Or pE >= 1 Or pT < 0 Or pT >= 1 Then
        HypgeomSample = CVErr(xlErrNum)
    Else
        n = Application.WorksheetFunction.RoundUp(Log(risk) / Log(1 - pT), 0)
        k = Application.WorksheetFunction.RoundUp(pE * n, 0)

        While Application.WorksheetFunction.HypGeom_Dist(k, n, Lt, P, True) > risk And n <= 20000
            n = n + 1
            k = Application.WorksheetFunction.RoundUp(pE * n, 0)
        Wend
        HypgeomSample = IIf(Application.WorksheetFunction.HypGeom_Dist(k, n, Lt, P, True) <= risk, n, CVErr(xlErrNA))
    End If
End Function


'for Attribute and MUS Sampling
Function BinomSample(risk As Double, pE As Double, pT As Double)
    Dim n As Double, k As Double
    If risk <= 0 Or risk >= 1 Or pE < 0 Or pE >= 1 Or pT <= 0 Or pT >= 1 Then
        BinomSample = CVErr(xlErrNum)
    Else
        n = Application.WorksheetFunction.RoundUp(Log(risk) / Log(1 - pT), 0)
        k = Application.WorksheetFunction.RoundUp(pE * n, 0)
        While Application.WorksheetFunction.BinomDist(k, n, pT, True) > risk And n <= 20000
            n = n + 1
            k = Application.WorksheetFunction.RoundUp(pE * n, 0)
        Wend
        BinomSample = IIf(Application.WorksheetFunction.BinomDist(k, n, pT, True) <= risk, n, CVErr(xlErrNA))
    End If
End Function


'for MUS Sampling
Function PoissonSample(risk As Double, pE As Double, pT As Double)
    Dim n As Double
    If risk <= 0 Or risk >= 1 Or pE < 0 Or pE >= 1 Or pT <= 0 Or pT >= 1 Then
        PoissonSample = CVErr(xlErrNum)
    Else
        n = Application.WorksheetFunction.RoundUp(-Log(risk) / pT, 0)
        While Application.WorksheetFunction.GammaDist(n * pT, 1 + pE * n, 1, True) < 1 - risk And n <= 20000
            n = n + 1
        Wend
        PoissonSample = IIf(Application.WorksheetFunction.GammaDist(n * pT, 1 + pE * n, 1, True) >= 1 - risk, n, CVErr(xlErrNA))
    End If
End Function