# VBA
For practice only.

'到第一个工作表
Sub GoToFirstSheet()

'选择最左边工作表
'插入工作表
On Error Resume Next
Sheets(1).Select             
Sheets.Add           

'提取每一个工作表名称
For Each sh In Sheets
k = k + 1                     
Cells(k, 1) = sh.Name
Next

'选择1,1单元格
 '删除首行
Rows("1:1").Select                
Selection.Delete Shift:=xlUp 

 '在A列左边插入1列
Columns("A:A").Insert       
   
                 
 '首列输入1-n
Dim i As Integer
    
n = WorksheetFunction.CountA([B:B])
    
i = 1

Do While i < n + 1

   Cells(i, 1) = i
  
   i = i + 1
    
Loop
    
    
MsgBox "共提取了" & n & "个工作表名称!"

End Sub

'插入空行

Option Explicit
Sub 插入空行()
On Error Resume Next
Dim rng As Range, a%, b%, c%, i&
Set rng = Application.InputBox("请选择要插入空行的单元格区域：", "请选择", Type:=8)
If Not rng Is Nothing Then
a = Val(InputBox("你想在目标区域中间隔多少行插入一次空白行？", "请输入", 1))
b = Val(InputBox("每次插入多少空白行数？", "请输入", 1))
c = Val(InputBox("插入空行的高度是？", "请输入", 15))
Application.ScreenUpdating = False
For i = 1 To rng.Rows.Count / a
rng.Cells((a + b) * i - b + 1, 1).Resize(b, rng.Columns.Count).Insert shift:=xlDown
rng.Cells((a + b) * i - b + 1, 1).Resize(b, 1).EntireRow.RowHeight = c
Next i
Application.ScreenUpdating = True
End If
Set rng = Nothing
End Sub

'合并单元格

Sub 合并单元格()
'
' 合并单元格 Macro
'
' 快捷键: Ctrl+s
'
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("B1:B2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("A1:B2").Select
    Selection.Copy
    Range("A1:B18").Select
    Range("A1:B2").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("A3:B18").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
End Sub



'选定区域空白单元格填充数字0
Sub s1()
 Dim rg As Range
 For Each rg In Range("a1:b7,d5:e9")  
   If rg = "" Then
     rg = 0
   End If
  Next rg
End Sub
