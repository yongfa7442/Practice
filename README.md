# VBA
For practice only.

Public n As Integer

Sub GoToFirstSheet()

On Error Resume Next
Sheets(1).Select             '选择最左边工作表
Sheets.Add                    '插入工作表

For Each sh In Sheets
k = k + 1                     '提取每一个工作表名称
Cells(k, 1) = sh.Name
Next

Rows("1:1").Select                '选择1,1单元格
    Selection.Delete Shift:=xlUp  '删除首行


   Columns("A:A").insert        '在A列左边插入1列
   
                    
MsgBox "共提取了" & k & "个工作表名称!"

End Sub
