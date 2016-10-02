'create a excel file
Function createExcel(strPath)
Const xlExcel7 = 39

' Spreadsheet file to be created.
strExcelPath = strPath

' Bind to Excel object.
Set objExcel = CreateObject("Excel.Application")
objExcel.Application.DisplayAlerts = False
' Code to create worksheet ...
Set objWorkBook = objExcel.Workbooks.Add
objExcel.SheetsInNewWorkbook = 1
Set objWorkSheet = objWorkBook.WorkSheets(1)
objWorkSheet.Name = "Execution Status"
objWorkSheet.Cells(1,1).Value = "Test Case No."
objWorkSheet.Cells(1,2).Value = "Test Case Execute"
objWorkSheet.Cells(2,1).Value = "TestCase 1"
objWorkSheet.Cells(3,1).Value = "TestCase 2"

Set objRange = objWorkSheet.Range("A1:D3")

objRange.Font.Name = "Arial"
objRange.Font.Size = "13"
objRange.Font.Bold = True
objRange.Interior.ColorIndex = "37"

REM 'objWorkSheet.Cells(3,1).Font.Size="20"
' Save the spreadsheet and close the workbook.
objExcel.ActiveWorkbook.SaveAs strExcelPath, xlExcel7
objExcel.ActiveWorkbook.Close
msgbox "Excel file created"
set objExcel = nothing
end function
call createExcel("c:\demo1.xls")