    '插入空行

    Sub Insert空行()

    Columns("A:A").insert        '在A列左边插入1列

    On Error Resume Next
    Dim rng As range, a%, b%, c%, i&
    Set rng = Application.InputBox("请选择要插入空行的单元格区域：", "请选择", Type:=8)
    If Not rng Is Nothing Then
    a = Val(InputBox("你想在目标区域中间隔多少行插入一次空白行？", "请输入", 1))
    b = Val(InputBox("每次插入多少空白行数？", "请输入", 1))
    c = Val(InputBox("插入空行的高度是？", "请输入", 15))
    Application.ScreenUpdating = False
    For i = 1 To rng.Rows.Count / a
    rng.Cells((a + b) * i - b + 1, 1).Resize(b, rng.Columns.Count).insert Shift:=xlDown
    rng.Cells((a + b) * i - b + 1, 1).Resize(b, 1).EntireRow.RowHeight = c
    Next i
    Application.ScreenUpdating = True
    End If
    Set rng = Nothing
    End Sub
    git config --global color.ui true
