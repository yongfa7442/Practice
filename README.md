Sub Data_()
Dim i%, a%, b%, c%, d%

 b = 3: c = 2: d = 8
Do While b < 6
a = 1
    For i = 1 To n
        Cells(a, b) = Worksheets(i + 1).Cells(3, c)
        Cells(a, d) = Worksheets(i + 1).Cells(5, c)
     a = a + 2
    Next
    b = b + 1
    c = c + 1
    d = d + 1
Loop


End Sub
