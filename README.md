       Sub Data_()
    Dim i%, a%, b%, c%
     b = 3: c = 2
    Do While b < 6
        a = 1
        For i = 1 To n
            Cells(a, b) = Worksheets(i + 1).Cells(3, c)
            Cells(a, b + 5) = Worksheets(i + 1).Cells(5, c)
            a = a + 2
        Next
        c = c + 1
        b = b + 1
    Loop

    'For b = 6 To 8 Step 1
       ' For i = 1 To n
     '       Cells(a, b) = Worksheets(i + 1).Cells(3, c)
      '      Cells(a, b + 5) = Worksheets(i + 1).Cells(5, c)
     '       a = a + 1
       ' Next
    'Next
