Function HypgeomUpperLimit(risk As Double, k As Double, n As Double, P As Double)
    Dim pT As Double, Lt As Double

    If P <= 0 Or k < 0 Or n <= 0 Or k > n Then
        HypgeomUpperLimit = CVErr(xlErrNum)
    Else
        pT = 0.25
        Lt = pT * P

        'Min value of pT such that Hypgeom_Dist(k, n, Lt, P) >= risk
        
        While Application.WorksheetFunction.HypGeom_Dist(k, n, Lt, P, True) < risk
            pT = pT - 0.001
            Lt = pT * P
        Wend
        HypgeomUpperLimit = pT
    End If
End Function

'pT as optional argument
Function HypgeomUpperLimit(risk As Double, k As Double, n As Double, P As Double, Optional pT As Double=0.25)
    Dim Lt As Double

    If P <= 0 Or k < 0 Or n <= 0 Or k > n Then
        HypgeomUpperLimit = CVErr(xlErrNum)
    Else
        Lt = pT * P

        'Min value of pT such that Hypgeom_Dist(k, n, Lt, P) >= risk
        
        While Application.WorksheetFunction.HypGeom_Dist(k, n, Lt, P, True) < risk
            pT = pT - 0.001
            Lt = pT * P
        Wend
        HypgeomUpperLimit = pT
    End If
End Function

Function BinomUpperLimit_Attr(risk As Double, k As Double, n As Double)
    If risk <= 0 Or risk >= 1 Or k > n Or k < 0 Or n <= 0 Then
        BinomUpperLimit_Attr = CVErr(xlErrNum)
    Else
        BinomUpperLimit_Attr = Application.WorksheetFunction.Beta_Inv(1-risk, 1+k, n-k)
    End If 
End Function

Function BinomUpperLimit_MUS(risk As Double, k As Double, n As Double)
    If risk <= 0 Or risk >= 1 Or k > n Or k < 0 Or n <= 0 Then
        BinomUpperLimit_MUS = CVErr(xlErrNum)
    ElseIf k = 0 Then
        BinomUpperLimit_MUS = Application.WorksheetFunction.Beta_Inv(1-risk, 1+k, n-k)
    ElseIf k > 0 Then
        BinomUpperLimit_MUS = Application.WorksheetFunction.Beta_Inv(1-risk, 1+k, n-k) - Application.WorksheetFunction.Beta_Inv(1-risk, 1+(k-1), n-(k-1))
    End If 
End Function

'PoissonUpperLimit equals the Reliability Factor Increment
Function PoissonUpperLimit(risk As Double, k As Double)
    If risk <= 0 Or risk >= 1 Or k <= 0 Then
        PoissonUpperLimit = CVErr(xlErrNum)
    Else
        PoissonUpperLimit = Application.WorksheetFunction.Gamma_Inv(1-risk, 1+k, 1) - Application.WorksheetFunction.Gamma_Inv(1-risk, 1+(k-1), 1)
    End If 
End Function

Function ReliabilityFactor(risk As Double, k As Double)
    If risk <= 0 Or risk >= 1 Or k < 0 Then
        ReliabilityFactor = CVErr(xlErrNum)
    Else
        ReliabilityFactor = Application.WorksheetFunction.Gamma_Inv(1-risk, 1+k, 1)
    End If 
End Function