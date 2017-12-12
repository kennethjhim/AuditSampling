'alternative to HYPGEOM.DIST() in Excel 2010 and up
Function dblHypgeom_Dist(k As Double, n As Double, pT As Double, P As Double)
    Dim Lt as Double
    Lt = Application.WorksheetFunction.RoundUp(pT * P, 0)
    dblHypgeom_Dist = Application.WorksheetFunction.HypGeom_Dist(k, n, Lt, P, True)
End Function

'for Excel 2007 and below
Function HGD(k As Double, n As Double, pT As Double, P As Double)
    Dim Lt As Double, Hg As Double, i As Double
    Lt = Application.WorksheetFunction.RoundUp(pT * P, 0)
    Hg = 0
    i = 0 
    While i < k
        Hg = Hg + Application.WorksheetFunction.HypGeomDist(i, n, Lt, P)
        i = i + 1
    Wend
    HGD = Hg
End Function