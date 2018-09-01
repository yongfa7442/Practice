    Sub TEST3()
    ' Merge
    ' 合并单元格
     Dim i%, a%
      FOR i =1 to 2 * n STEP 2   
        FOR a = 2 to 13 STEP 1
            If (a <> 6 And a <> 7 And a <> 11 And a <> 12) Then
                Range(Cells(i, a), Cells(i + 1, a)).Merge
            End If
        NEXT
      NEXT
    End Sub
