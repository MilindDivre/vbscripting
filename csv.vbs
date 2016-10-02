filename="C:\VB Script Training\CSVFolder\ENQ48.csv"
'filename="C:\VB Script Training\CSVFolder\test2.csv"

call changeColumnName(filename,"Pin","1234")
function changeColumnName(filename,SrcData,strReplaceWith)
	Set objExcel = CreateObject("Excel.Application")
	'objExcel.Visible = True
	Set objwrkbook = objExcel.Workbooks.Open(filename)
	Set objWorksheet = objwrkbook.Worksheets(1)
	i = 1
	do until  objWorksheet.Cells(1, i)= ""
		data=objExcel.Cells(1,i)
		'msgbox data
		val=strcomp(data,SrcData)
		if val = 0 then
			msgbox "match found"
			rowVal = objExcel.Cells(2,i).value
			rSearch = Replace(rowVal,rowVal,strReplaceWith)
			'msgbox rSearch
			objExcel.Cells(2,i).value = rSearch
	 end if
	i = i + 1
	loop
	objExcel.DisplayAlerts = False
	'msgbox filename
	objExcel.Workbooks(1).SaveAs filename, 6 
	objwrkbook.Saved = True
	objExcel.ScreenUpdating = True

	set objwrkbook = nothing
	set objWorksheet = nothing
	objExcel.quit
End Function 	
	
	'read using simple text file
	
	REM set objCSV = CreateObject("Scripting.FileSystemObject")
	REM set objreadFile = objCSV.opentextfile(filename)
	REM Do Until objreadFile.AtEndOfStream
		REM line=objreadFile.ReadLine
		REM msgbox LTrim(line)
		REM records = split(line,",")
	REM Loop
	REM objreadFile.Close
	REM for each i in records
		REM MsgBox I
	REM next
	
	'read using database technique
	
	REM On Error Resume Next
REM Const adOpenStatic = 3
REM Const adLockOptimistic = 3
REM Const adCmdText = &H0001

REM Set objConnection = CreateObject("ADODB.Connection")
REM Set objRecordSet = CreateObject("ADODB.Recordset")

REM strPathtoTextFile = "C:\VB Script Training\CSVFolder\"

REM objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          REM "Data Source=" & strPathtoTextFile & ";" & _
          REM "Extended Properties=""text;HDR=YES;FMT=Delimited"""

REM objRecordset.Open "SELECT * FROM test2.csv ", _
          REM objConnection, adOpenStatic, adLockOptimistic, adCmdText

REM Do Until objRecordset.EOF
    REM msgbox "Roll No " & objRecordset.Fields.Item("RNO")
    REM msgbox "name: " & objRecordset.Fields.Item("NAME")
    REM msgbox "grade: " & objRecordset.Fields.Item("grade")   
    REM objRecordset.MoveNext
REM Loop


	