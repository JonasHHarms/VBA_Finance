Public Function get_Vega(S As Double, k As Double, r As Double, vol As Double, dy As Double, T As Double) As Double
    Dim d_1 As Double
    Dim dxNd_1 As Double
    
        d_1 = (Log(S / k) + (r - dy + 0.5 * (vol ^ 2)) * T) / (vol * Sqr(T))
        dxNd_1 = standard_normal_pdf(d_1)
        get_Vega = S * Exp(-dy * T) * Sqr(T) * dxNd_1
End Function
