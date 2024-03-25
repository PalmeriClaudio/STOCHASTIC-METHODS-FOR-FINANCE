Attribute VB_Name = "Module2"
Function Binomial(S As Double, K As Double, r As Double, T As Double, vol As Double, n As Integer) As Double
    
    Dim dt As Double
    Dim u As Double
    Dim d As Double
    Dim p As Double
    Dim i As Integer
    Dim j As Integer
    Dim Sj As Double
   
    dt = T / n
    u = Exp(vol * Sqr(dt))
    d = 1 / u
    p = (Exp(r * dt) - d) / (u - d)
   
    ReDim v(0 To n, 0 To n) As Double
    
    'Calculate stock prices at maturity
    For j = 0 To n
        Sj = S * (u ^ (n - j)) * (d ^ j)
        v(n, j) = WorksheetFunction.Max(Sj - K, 0)
    Next j
   
    'Calculate option prices at earlier times
    For i = n - 1 To 0 Step -1
        For j = 0 To i
            Sj = S * (u ^ (i - j)) * (d ^ j)
            v(i, j) = Exp(-r * dt) * (p * v(i + 1, j) + (1 - p) * v(i + 1, j + 1))
        Next j
    Next i
   
    'Return option price
    Binomial = v(0, 0)
End Function
Function BlackScholes(S As Double, K As Double, r As Double, T As Double, vol As Double) As Double
    
    Dim d1 As Double
    Dim d2 As Double
   
    d1 = (Log(S / K) + (r + (vol ^ 2) / 2) * T) / (vol * Sqr(T))
    d2 = d1 - vol * Sqr(T)
   
    BlackScholes = S * WorksheetFunction.NormSDist(d1) - K * Exp(-r * T) * WorksheetFunction.NormSDist(d2)
End Function
Function LeisenReimer(S As Double, K As Double, r As Double, T As Double, vol As Double, n As Integer) As Double
    Dim d1 As Double
    Dim d2 As Double
    Dim hd1 As Double
    Dim hd2 As Double
    Dim dt As Double
    Dim p As Double
    Dim u As Double
    Dim d As Double
    Dim i As Integer
    Dim j As Integer
    Dim Sj As Double
    
    d1 = (Log(S / K) + (r + vol ^ 2 / 2) * T) / (vol * Sqr(T))
    d2 = d1 - vol * Sqr(T)
    hd1 = 0.5 + Sgn(d1) * (0.25 - 0.25 * Exp(-(d1 / (n + 1 / 3 + 0.1 / (n + 1))) ^ 2 * (n + 1 / 6))) ^ 0.5
    hd2 = 0.5 + Sgn(d2) * (0.25 - 0.25 * Exp(-(d2 / (n + 1 / 3 + 0.1 / (n + 1))) ^ 2 * (n + 1 / 6))) ^ 0.5
    dt = T / n
    p = hd2
    u = Exp(r * dt) * hd1 / hd2
    d = Exp(r * dt) * (1 - hd1) / (1 - hd2)
    
    ReDim v(0 To n, 0 To n) As Double
   
    'Calculate stock prices at maturity
    For j = 0 To n
        Sj = S * (u ^ (n - j)) * (d ^ j)
        v(n, j) = WorksheetFunction.Max(Sj - K, 0)
    Next j
   
    'Calculate option prices at earlier times
    For i = n - 1 To 0 Step -1
        For j = 0 To i
            Sj = S * (u ^ (i - j)) * (d ^ j)
            v(i, j) = Exp(-r * dt) * (p * v(i + 1, j) + (1 - p) * v(i + 1, j + 1))
        Next j
    Next i
   
    'Return option price
    LeisenReimer = v(0, 0)
End Function
