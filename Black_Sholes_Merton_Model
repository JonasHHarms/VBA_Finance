Public Function Black_Scholes_Merton(CP As String, S As Double, k As Double, r As Double, v As Double, dy As Double, T As Double) As Double
    Dim d_1  As Double
    Dim d_2  As Double
    Dim Premium As Double
    Dim Nd_1 As Double
    Dim Nd_2 As Double
    Dim sk As Double
    
     
    sk = S / k
    d_1 = (Log(sk) + (r - dy + 0.5 * (v ^ 2)) * T) / (v * Sqr(T))
    d_2 = d_1 - v * Sqr(T)
    
    
    If CP = "C" Then
        Nd_1 = cumulative_st_ndist(d_1)
        Nd_2 = cumulative_st_ndist(d_2)
        Premium = S * Nd_1 * Exp(-dy * T) - k * Nd_2 * Exp(-r * T)
    ElseIf CP = "P" Then
        Nd_1 = cumulative_st_ndist(-d_1)
        Nd_2 = cumulative_st_ndist(-d_2)
        Premium = -S * Nd_1 * Exp(-dy * T) + k * Nd_2 * Exp(-r * T)
    Else
        MsgBox "Unexpected option Type, expected C or P. Check in Boxes_references"
        End
    End If
       
    Black_Scholes_Merton = Premium
       
End Function
