  Public Function EasterDate(Y As Integer) As Date
    Dim a As Integer: a = Y Mod 19
    Dim b As Integer: b = Y \ 100
    Dim c As Integer: c = Y Mod 100
    Dim d As Integer: d = b \ 4
    Dim e As Integer: e = b Mod 4
    Dim F As Integer: F = (b + 8) \ 25
    Dim g As Integer: g = (b - F + 1) \ 3
    Dim h As Integer: h = (19 * a + b - d - g + 15) Mod 30
    Dim i As Integer: i = c \ 4
    Dim k As Integer: k = c Mod 4
    Dim l As Integer: l = (32 + 2 * e + 2 * i - h - k) Mod 7
    Dim m As Integer: m = (a + 11 * h + 22 * l) \ 451
    Dim month As Integer: month = (h + l - 7 * m + 114) \ 31
    Dim day As Integer: day = ((h + l - 7 * m + 114) Mod 31) + 1
    EasterDate = DateSerial(Y, month, day)
End Function
