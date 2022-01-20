Dim ws
Dim Counter
Dim objExcel

Counter = 0
Set objExcel =  GetObject(, "Excel.Application")

For Each ws In objExcel.ActiveWorkbook.Worksheets
    objExcel.ActiveCell.Offset(Counter, 0).Value = ws.Name
    Counter = Counter + 1
Next



