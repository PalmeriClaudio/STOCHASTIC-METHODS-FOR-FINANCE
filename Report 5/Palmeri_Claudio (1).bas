Attribute VB_Name = "Module1"
Function DELTAGREEK(S As Double, K As Double, r As Double, T As Double, vol As Double) As Double
    Dim d1 As Double
    d1 = (Log(S / K) + (r + (vol ^ 2) / 2) * T) / (vol * Sqr(T))
    DELTAGREEK = WorksheetFunction.NormSDist(d1)
End Function
Function GAMMAGREEK(S As Double, K As Double, r As Double, T As Double, vol As Double) As Double
    Dim d1 As Double
    d1 = (Log(S / K) + (r + (vol ^ 2) / 2) * T) / (vol * Sqr(T))
    GAMMAGREEK = (Exp(-d1 * d1 / 2)) / (S * vol * Sqr(2 * WorksheetFunction.Pi()) * Sqr(T))
End Function
Function RHOGREEK(S As Double, K As Double, r As Double, T As Double, vol As Double) As Double
    Dim d1 As Double
    Dim d2 As Double
    d1 = (Log(S / K) + (r + (vol ^ 2) / 2) * T) / (vol * Sqr(T))
    d2 = d1 - vol * Sqr(T)
    RHOGREEK = K * T * Exp(-r * T) * WorksheetFunction.NormSDist(d2)
End Function
Function THETAGREEK(S As Double, K As Double, r As Double, T As Double, vol As Double) As Double
    Dim d1 As Double
    Dim d2 As Double
    d1 = (Log(S / K) + (r + (vol ^ 2) / 2) * T) / (vol * Sqr(T))
    d2 = d1 - vol * Sqr(T)
    THETAGREEK = (-S * Exp(-d1 * d1 / 2)) / (2 * Sqr(T) * Sqr(2 * WorksheetFunction.Pi())) - r * K * Exp(-r * T) * WorksheetFunction.NormSDist(d2)
End Function
Function VEGAGREEK(S As Double, K As Double, r As Double, T As Double, vol As Double) As Double
    Dim d1 As Double
    d1 = (Log(S / K) + (r + (vol ^ 2) / 2) * T) / (vol * Sqr(T))
    VEGAGREEK = S * Exp(-d1 * d1 / 2) * Sqr(T) / Sqr(2 * WorksheetFunction.Pi())

End Function

