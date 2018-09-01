    Public n%
    
    Sub p4p6()
	    n = Worksheets.Count
	    Call TEST1
	    Call TEST2
	    Call TEST4
	    Call Data_
    End Sub
    
    Sub TEST1()  'Go to fistsheet
		On Error Resume Next
		Sheets(1).Select             '选择最左边工作表
		Sheets.Add                    '插入工作表
		For Each sh In Sheets
			k = k + 1                     '提取每一个工作表名称
			Cells(k, 1) = sh.Name
		Next
		Rows("1:1").Select                '选择1,1单元格
		Selection.Delete Shift:=xlUp  '删除首行
    End Sub
