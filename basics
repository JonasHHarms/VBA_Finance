Public Function Max(ParamArray p()) As Variant
    Dim i As Long
    Dim v As Variant

    v = p(LBound(p))
    For i = LBound(p) + 1 To UBound(p)
      If v < p(i) Then
         v = p(i)
      End If
    Next
    Max = v
  End Function
  

  Public Function Min(ParamArray p()) As Variant
    Dim i As Long
    Dim v As Variant

    v = p(LBound(p))
    For i = LBound(p) + 1 To UBound(p)
      If v > p(i) Then
         v = p(i)
      End If
    Next
    Min = v
  End Function
