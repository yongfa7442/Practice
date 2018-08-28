'copy & paste
Sub copy()


'首列输入日期
    Dim i As Integer
    
    n = WorksheetFunction.CountA([B:B])
    
    i = 1
    Do While i < 2 * n + 1
    
        Cells(i, 1) = Format(Date - 2)
        
          Cells(i, 1).NumberFormat = "m-d"
          
        i = i + 1
    Loop

End Sub
