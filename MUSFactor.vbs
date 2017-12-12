Function MUSFactor(risk As Double, e As Double)
    Dim F1 As Double, F As Double
    Dim i As Integer
    If risk <= 0 Or risk >= 1 Or e < 0 Or e >= 1 Then
        MUSFactor = CVErr(xlErrNum)
    Else
        F = Application.WorksheetFunction.GammaInv(1 - risk, 1, 1)
        If e = 0 Then
            MUSFactor = F
        Else
            F1 = 0
            i = 0
            While Abs(F1 - F) > 0.000001 And i <= 1000
                F1 = F
                F = Application.WorksheetFunction.GammaInv(1 - risk, 1 + e * F1, 1)
                i = i + 1
            Wend
            MUSFactor = IIf(Abs(F1 - F) <= 0.000001, F, CVErr(xlErrNum))
        End If
    End If
End Function

