set objExcel = CreateObject("Excel.Application")
objExcel.Application.DisplayAlerts = False

set objWorkbook=objExcel.workbooks.add()
objExcel.cells(1,1).value = "Test value"
objExcel.cells(1,2).value = "Test data"

objWorkbook.Saveas "c:\testXLS.xls"
objWorkbook.Close
objExcel.workbooks.close
objExcel.quit

REM Dim objexcel, objWorkbook, objDriverSheet, columncount, rowcount
REM set objexcel = Createobject("Excel.Application")
REM Set objWorkbook = objExcel.WorkBooks.Open("c:\testXLS.xls")
REM Set objDriverSheet = objWorkbook.Worksheets(1)
REM columncount = objDriverSheet.usedrange.columns.count ' gives number of column utilised
REM msgbox columncount
REM rowcount = objDriverSheet.usedrange.rows.count'gives number of row utilised
REM msgbox  rowcount
REM for i = 1 to columncount
	REM columnname = objDriversheet.cells(i,1)
	REM msgbox columnname
		REM for j = 1 to rowcount
			REM fieldvalue = objdriversheet.cells(j,i)
			REM msgbox fieldvalue
		REM next
REM next
REM Set objWorkbook = nothing
REM objexcel.quit