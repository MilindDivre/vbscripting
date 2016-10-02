set objExcel  = CreateObject("Excel.Application")
set objWorkbook = objExcel.workbooks.open("C:\VB Script Training\test.xlsx")
set objdriverSheet = objWorkbook.worksheets(1)
columncount = objDriverSheet.usedrange.columns.count ' gives number of column utilised
msgbox columncount
rowcount = objDriverSheet.usedrange.rows.count'gives number of row utilised
msgbox  rowcount
Set objWorkbook =  nothing
set objExcel = nothing