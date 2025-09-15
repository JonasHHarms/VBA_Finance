Public Function Newton_Raphson(CP As String, S As Double, k As Double, r As Double, true_Price As Double, dy As Double, T As Double) As Double
    Dim dVol As Double
    Dim epsilon As Double
    Dim maxIter As Double
    Dim vol As Double
    Dim dxNd_1 As Double
    Dim d_1 As Double
    Dim Model_price As Double
    Dim Vega As Double
    Dim Value_minus As Double
    Dim i As Integer
    Dim dx As Double
    Dim floor As Double
    Dim cap As Double
    Dim Lsk As Double
    Dim rdy As Double
    Dim vmin As Double
    Dim vmax As Double
    
    
        
    dVol = 0.0001
    epsilon = 0.00001
    maxIter = 100
    vol = 0.2
    cap = 2
    floor = 0.01
    vmin = floor
    vmax = 1
    
    For i = 1 To maxIter
        If vol < floor Or vol > cap Then
           vol = (vmin + vmax) / 2
        End If

        Model_price = option_pricer(CP, S, k, r, vol, dy, T)
        Vega = get_Vega(S, k, r, vol, dy, T)
        
        If Model_price > true_Price Then
            vmax = vol
        Else
            vmin = vol
        End If
        
        If Abs(Vega) < epsilon Then Exit For
        
        If Abs(Model_price - true_Price) < epsilon Then Exit For
        
        If (vmax - vmin) < epsilon Then Exit For
        
        vol = vol - (Model_price - true_Price) / Vega
        
        If i > 50 Then
            MsgBox "Vola dreht ab, bitte mal die local variables in VBA prüfen"
        End If
        
    Next i
    
    
    
    Newton_Raphson = vol

End Function
