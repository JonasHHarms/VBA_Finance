Public Function quickspline(x As Double, xx() As Double, yy() As Double) As Double
    Dim i As Integer
    Dim gxx(0 To 1) As Double
    Dim ggxx(0 To 1) As Double
    Dim a As Double
    Dim b As Double

    i = Num
    a = (xx(i) - x) / (xx(i) - xx(i - 1))
    b = 1 - a
    Cy = 1 / 6 * (a ^ 3 - a) * (6 * (yy(i) - yy(i - 1)) - 2 * (gxx(i) + 2 * gxx(i - 1)) * (xx(i) - xx(i - 1)))
    dy = 1 / 6 * (b ^ 3 - b) * (2 * (2 * gxx(i) + gxx(i - 1)) * (xx(i) - xx(i - 1)) - 6 * (yy(i) - yy(i - 1)))

    quickspline = a * yy(i - 1) + b * yy(i) + Cy + dy
End Function


Public Function Spline(x As Double, xx() As Double, yy() As Double) As Double
' referenced from Kruger
    Dim i As Integer
    Dim j As Integer
    Dim Nmax As Integer
    Dim Num As Integer
    Dim gxx(0 To 1) As Double
    Dim ggxx(0 To 1) As Double
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    
    Nmax = UBound(xx())
    Num = 0
    
    If x < xx(0) Or x > xx(Nmax) Then
      If x < xx(0) Then Num = 1 Else Num = Nmax
      b = (yy(Num) - yy(Num - 1)) / (xx(Num) - xx(Num - 1))
      a = yy(Num) - b * xx(Num)
      SplineX3 = a + b * x
      Exit Function
    Else
      For i = 1 To Nmax
        If x <= xx(i) Then
          Num = i
          Exit For
        End If
      Next i
    End If
    
    For j = 0 To 1
      i = Num - 1 + j
      If i = 0 Or i = Nmax Then
        gxx(j) = 10 ^ 30
      ElseIf (yy(i + 1) - yy(i) = 0) Or (yy(i) - yy(i - 1) = 0) Then
        gxx(j) = 0
      ElseIf ((xx(i + 1) - xx(i)) / (yy(i + 1) - yy(i)) + (xx(i) - xx(i - 1)) / (yy(i) - yy(i - 1))) = 0 Then
        gxx(j) = 0
      ElseIf (yy(i + 1) - yy(i)) * (yy(i) - yy(i - 1)) < 0 Then
        gxx(j) = 0
      Else
        gxx(j) = 2 / ((xx(i + 1) - xx(i)) / (yy(i + 1) - yy(i)) + (xx(i) - xx(i - 1)) / (yy(i) - yy(i - 1)))
      End If
    Next j
    
    If Num = 1 Then
      gxx(0) = 3 / 2 * (yy(Num) - yy(Num - 1)) / (xx(Num) - xx(Num - 1)) - gxx(1) / 2
    End If
    If Num = Nmax Then
      gxx(1) = 3 / 2 * (yy(Num) - yy(Num - 1)) / (xx(Num) - xx(Num - 1)) - gxx(0) / 2
    End If
    
    ggxx(0) = -2 * (gxx(1) + 2 * gxx(0)) / (xx(Num) - xx(Num - 1)) + 6 * (yy(Num) - yy(Num - 1)) / (xx(Num) - xx(Num - 1)) ^ 2
    ggxx(1) = 2 * (2 * gxx(1) + gxx(0)) / (xx(Num) - xx(Num - 1)) - 6 * (yy(Num) - yy(Num - 1)) / (xx(Num) - xx(Num - 1)) ^ 2
    
    d = 1 / 6 * (ggxx(1) - ggxx(0)) / (xx(Num) - xx(Num - 1))
    c = 1 / 2 * (xx(Num) * ggxx(0) - xx(Num - 1) * ggxx(1)) / (xx(Num) - xx(Num - 1))
    b = (yy(Num) - yy(Num - 1) - c * (xx(Num) ^ 2 - xx(Num - 1) ^ 2) - d * (xx(Num) ^ 3 - xx(Num - 1) ^ 3)) / (xx(Num) - xx(Num - 1))
    a = yy(Num - 1) - b * xx(Num - 1) - c * xx(Num - 1) ^ 2 - d * xx(Num - 1) ^ 3
    
    Spline = a + b * x + c * x ^ 2 + d * x ^ 3
    
End Function
