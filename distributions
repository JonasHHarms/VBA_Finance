Public Function cumulative_st_ndist(x As Double) As Double
    'computes the cumulative standard normal distribution
    'constants from „Handbook of Mathematical Functions“ von Abramowitz & Stegun (Tabelle 7.1.26)
    Const a1 As Double = 0.31938153
    Const a2 As Double = -0.356563782
    Const a3 As Double = 1.781477937
    Const a4 As Double = -1.821255978
    Const a5 As Double = 1.330274429
    Const Pi As Double = 3.14159265358979
    Dim l As Double
    Dim k As Double
    Dim w As Double
    
    l = Abs(x)
    k = 1 / (1 + 0.2316419 * l)
    
    w = 1 - 1 / Sqr(2 * Pi) * Exp(-l * l / 2) * (a1 * k + a2 * k ^ 2 + a3 * k ^ 3 + a4 * k ^ 4 + a5 * k ^ 5)
    
    If x < 0 Then
        cumulative_st_ndist = 1 - w
    Else
        cumulative_st_ndist = w
    End If
End Function


Public Function standard_normal_pdf(x As Double) As Double
    Const Pi As Double = 3.14159265358979
    standard_normal_pdf = Exp(-x ^ 2 / 2) / Sqr(2 * Pi)
End Function
