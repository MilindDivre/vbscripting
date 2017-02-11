Function GetDicObjectFromExternalSource(myXlsFile, mySheet, PrimaryKey)

	' Function :  GetDicObjectFromExternalSource
	' This function reads data from an Excel sheet without using MS-Office
	'
	' Arguments:
	' myXlsFile   [string]   The path and file name of the Excel file
	' mySheet     [string]   The name of the worksheet used (e.g. "Sheet1")
	
	' Returns:
	' The values of first row read from the Excel sheet are returned in a dictionary object
	msgbox "sheetname->"&mySheet
	Dim arrData(), arrExecDetails(), i, j
	Dim objExcel, objRS
	Dim strHeader, strRange
	
	Const adOpenForwardOnly = 0
	Const adOpenKeyset = 1
	Const adOpenDynamic = 2
	Const adOpenStatic = 3
	
	' Define header parameter string for Excel object
	strHeader = "HDR=YES;" ' This means the first row is header
	
	On Error Resume Next
	
	' Open the object for the Excel file
	Set objExcel = CreateObject("ADODB.Connection")
	'	    objExcel.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
	'	                  myXlsFile & ";Extended Properties=""Excel 12.0;IMEX=1;" & _
	'	                  strHeader & """"
	sconnect = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & myXlsFile & ";HDR=Yes';"  
	objExcel.Open sconnect           
	
	If Err.Number <> 0 Then
		Reporter.ReportEvent micFail, "GetDicObjectFromExternalSource", Err.Description
		PrintLog "GetDicObjectFromExternalSource-->"& Err.Description
		GetDicObjectFromExternalSource = Empty
		Exit Function
	End If
	msgbox "sads"
	' Open a recordset object for the sheet and range
	Set objRS = CreateObject("ADODB.Recordset")
	strRange = mySheet & "$"
	PrimaryKey= cstr(PrimaryKey)
	sql = "Select * from [" & strRange & "] WHERE PrimaryKey='" & PrimaryKey & "'"
	msgbox sql
	'objRS.Open "Select * from [" & strRange & "] WHERE PrimaryKey = '"& PrimaryKey &"' ", objExcel, adOpenStatic
	'objRS.Open "Select * from [" & strRange & "] WHERE PrimaryKey="&PrimaryKey, objExcel, adOpenStatic
	'objRS.Open "Select * from [" & strRange & "] Where PrimaryKey = "& PrimaryKey &"", objExcel, adOpenStatic
	objRS.Open sql, objExcel, adOpenStatic
	If Err.Number <> 0 Then
		msgbox Err.Description
		Reporter.ReportEvent micFail, "GetDicObjectFromExternalSource", Err.Description
		PrintLog "GetDicObjectFromExternalSource-->"& Err.Description
		GetDicObjectFromExternalSource = Empty
		Exit Function
	End If
	
	On Error Goto 0
	
	Set objDataRow = CreateObject("Scripting.Dictionary")
	
	strFlag = False
	
	' Read the data from the Excel sheet
	Do Until objRS.EOF
		strFlag = True
		' Stop reading when an empty row is encountered in the Excel sheet
		If IsNull(objRS.Fields(0).Value) Or Trim(objRS.Fields(0).Value) = "" Then Exit Do
		
		' IsNull test credits: Adriaan Westra
		For j = 0 To objRS.Fields.Count - 1
			'Print Trim(objRS.Fields(j).Name) & "--->" & Trim(objRS.Fields(j).Value)
			sCellName = Trim(objRS.Fields(j).Name)
			sCellValue = Trim(objRS.Fields(j).Value)
			objDataRow.Item(sCellName) = sCellValue
		Next
		' Move to the next row
		objRS.MoveNext
		Exit Do
	Loop
	
	' Close the file and release the objects
	objRS.Close
	objExcel.Close
	Set objRS = Nothing
	Set objExcel = Nothing
	
	' Return the results
	If strFlag= True Then
		Set GetDicObjectFromExternalSource = objDataRow
		msgbox "sad"
		for each item in objDataRow
			msgbox item & "->"&objDataRow.item(item)
		next
		Set objDataRow = Nothing
	Else
		GetDicObjectFromExternalSource = Empty
	End If
	
End Function

pri_key="test_1"
msgbox pri_key
call GetDicObjectFromExternalSource ("C:\Users\divrem\Desktop\VB Script Training\readData.xlsx", "Sheet1", pri_key)
'for each item in GetData
'	 msgbox item & "->" &GetData.item(item)
'next