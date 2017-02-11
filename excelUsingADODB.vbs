
'sFileName= xls file name with path
'Sheetname = table name
'VRstatmt = where statment for query
'Colname=Column name to fetch value from
'==============================================
Public Function QryXls_GetData( sFileName, SheetName, VRstatmt,Colname )
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = "&H0001"

sql_text="Select * FROM [" & SheetName & "$]" & VRstatmt
 '1) Create an ADODB connection and recordset
Set objConnection = CreateObject("ADODB.Connection")
msgbox sql_text
'2) Open connection'
objConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & sFileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;"";"

'3) Create a Recordset'
Set objRecordSet = CreateObject("ADODB.Recordset")

'4) Execute SQL and store results in reocrdset'
objRecordset.Open sql_text , objConnection, adOpenStatic, adLockOptimistic, adCmdText

'5) Read all fields data   and store in a array'
'For Multiple records
'=================
ReDim  SQLExpectedData(objRecordset.recordcount -1)
ReDim  orderkind(objRecordset.recordcount -1)
For i=0  to objRecordset.recordcount -1  'objRecordset.fields.item(1).properties.count
    SQLExpectedData(i)= objRecordset.fields(Colname)
	orderkind(i)=objRecordset.fields("order_kind")
    objRecordset.movenext
Next

msgbox ubound(SQLExpectedData)
for i =0 to ubound(SQLExpectedData)
msgbox SQLExpectedData(i) & orderkind(i)
next




'6) Close and Discard all variables '
objRecordset.Close
objConnection.Close
End Function

call QryXls_GetData("test.xlsx","Sheet1","where run_status='yes'","id")