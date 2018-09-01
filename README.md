'插入空行
Sub TEST2()   'INSERT

Columns("A:A").Insert '在A列左边插入1列

Dim a%

a = 2

Do While a < 2 * n + 1
        Rows(a).Insert  '插入a行
        a = a + 2
  Loop
    
End Sub
