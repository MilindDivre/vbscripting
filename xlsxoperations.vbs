REM set objExcel = CreateObject("Excel.Application")
REM objExcel.Application.DisplayAlerts = False
REM objExcel.workbooks.add
REM objExcel.cells(1,1).value = "Test value"
REM objExcel.cells(1,2).value = "Test data"
REM objExcel.cells(1,2).value = "Test data"
REM objExcel.cells(1,2).value = "Test data"
REM objExcel.cells(1,2).value = "Test data"
REM objExcel.ActiveWorkbook.Saveas "c:\testXLS.xlsx"
REM objExcel.ActiveWorkbook.Close
REM objExcel.workbooks.close
REM objExcel.quit
REM msgbox  "done!!"
REM set objExcel = nothing 

set objExcel  = CreateObject("Excel.Application")
set objWorkbook = objExcel.workbooks.open("c:\testXLS.xlsx")
set objdriverSheet = objWorkbook.worksheets(1)
columncount = objDriverSheet.usedrange.columns.count ' gives number of column utilised
msgbox columncount
rowcount = objDriverSheet.usedrange.rows.count'gives number of row utilised
msgbox  rowcount
for i = 1 to columncount
	columnname = objDriversheet.cells(i,1)
	'msgbox "Column-"&columnname &"value of i"&i
		for j = 1 to rowcount
			fieldvalue = objdriversheet.cells(j,i)
			msgbox "row-"&fieldvalue
			if fieldvalue = "Pune" then
				objdriversheet.cells(j,i).Interior.ColorIndex = 3
			end if
		next
	
next
objExcel.ActiveWorkbook.Save
	
	objExcel.ActiveWorkbook.Close


	objExcel.Application.Quit
Set objWorkbook =nothing
objexcel.quit