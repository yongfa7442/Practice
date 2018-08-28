# VBA
For practice only.

'选定区域空白单元格填充数字0
Sub s1()
 Dim rg As Range
 For Each rg In Range("a1:b7,d5:e9")  
   If rg = "" Then
     rg = 0
   End If
  Next rg
End Sub
