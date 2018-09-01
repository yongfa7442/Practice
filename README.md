    Sub TEST4() 
    '首列输入日期
    Dim i As Integer        
       FOR i =1 TO 2 * n + 1 STEP 1
            Cells(i, 1) = Format(Date - 1)  
            Cells(i, 1).NumberFormat = "m-d"   
       NEXT
    End Sub
