Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set xlVbscript = objExcel.WorkBooks.Open("C:\ExcelFile.xlsx")
Const xlDown = -4121
Const xlToRight = -4161

''''''''''''''It Should be in two mwthod'''''''''''''''''''



''''''''''''''''''''Here its a 1st Method(If There is No Empty Rows And Column Using This method)''''''''''''''''''
ExcelRow = xlVbscript.Sheets(1).Range("A1").End(xlDown).Row
MsgBox(ExcelRow)

ExcelColumn = xlVbscript.Sheets(1).Range("A1").End(xlToRight).Column
MsgBox(ExcelColumn)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




''''''''''''''''''''''Here its a 2nd Method''''''''''''''''
xlColRange = xlVbscript.Sheets(1).UsedRange.Columns.Count
MsgBox(xlColRange)

xlRowRange = xlVbscript.Sheets(1).UsedRange.Rows.Count
MsgBox(xlRowRange)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
xlVbscript.save
xlVbscript.Close

objExcel.Quit
set objExcel=nothing