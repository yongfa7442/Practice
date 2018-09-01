    Sub TEST4() 'copy

    'copy & paste
    '首列输入日期
 
    Dim i As Integer
        
        i = 1
        
        Do While i < 2 * n + 1
    
        Cells(i, 1) = Format(Date - 1)
        
          Cells(i, 1).NumberFormat = "m-d"
          
        i = i + 1
    Loop

    End Sub
