' 合并单元格
Option Explicit

Sub Merge()

n = WorksheetFunction.CountA([B:B])

  Dim i As Integer
 
  i = 1
 
 Do While i <= 2 * n
  
    Cells(i + 1, 1).copy Cells(i, 1)
   
    range(Cells(i, 2), Cells(i + 1, 2)).Merge
   
    i = i + 2
    
  Loop
    
End Sub
