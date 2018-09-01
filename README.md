    Sub TEST3()
   ' Merge
   ' 合并单元格

      Dim i%, a%
      i = 1

        Do While i <= 2 * n
            a = 2
   
       Do While a <= 13
       If (a <> 6 And a <> 7 And a <> 11 And a <> 12) Then

         Range(Cells(i, a), Cells(i + 1, a)).Merge
        End If
         a = a + 1
      Loop

        i = i + 2

       Loop

    End Sub
