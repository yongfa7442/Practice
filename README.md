# VBA
For practice only.

Public n As Integer

Sub GoToFirstSheet()

On Error Resume Next
Sheets(1).Select             'Ñ¡Ôñ×î×ó±ß¹¤×÷±í
Sheets.Add                    '²åÈë¹¤×÷±í

For Each sh In Sheets
k = k + 1                     'ÌáÈ¡Ã¿Ò»¸ö¹¤×÷±íÃû³Æ
Cells(k, 1) = sh.Name
Next

Rows("1:1").Select                'Ñ¡Ôñ1,1µ¥Ôª¸ñ
    Selection.Delete Shift:=xlUp  'É¾³ýÊ×ÐÐ


     Columns("A:A").insert        'ÔÚAÁÐ×ó±ß²åÈë1ÁÐ
   
                    
MsgBox "¹²ÌáÈ¡ÁË" & k & "¸ö¹¤×÷±íÃû³Æ!"

End Sub

    
  
