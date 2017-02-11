REM Const adOpenStatic = 3 ' adding somthing to branch
REM Const adLockOptimistic = 3

REM Set objConnection = CreateObject("ADODB.Connection")
REM Set objRecordSet = CreateObject("ADODB.Recordset")

REM objConnection.Open _
    REM "Provider = Microsoft.Jet.OLEDB.4.0; " & _
        REM "Data Source = C:\VB Script Training\test.accdb"

REM objRecordSet.Open "INSERT INTO Employee (Emp_Name, Designation, Salary)" &  _
    REM "VALUES ('Test2', 'Manager', '20000')", _
        REM objConnection, adOpenStatic, adLockOptimistic


REM Dim connStr, objConn, getNames
REM '''''''''''''''''''''''''''''''''''''
REM 'Define the driver and data source
REM 'Access 2007, 2010, 2013 ACCDB:
REM 'Provider=Microsoft.ACE.OLEDB.12.0
REM 'Access 2000, 2002-2003 MDB:
REM 'Provider=Microsoft.Jet.OLEDB.4.0
REM ''''''''''''''''''''''''''''''''''''''
REM connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\VB Script Training\test.accdb"
 
REM 'Define object type
REM Set objConn = CreateObject("ADODB.Connection")
 
REM 'Open Connection
REM objConn.open connStr
 
REM 'Define recordset and SQL query
REM Set rs = objConn.execute("SELECT Emp_Name FROM Employee")
 
REM 'While loop, loops through all available results
REM DO WHILE NOT rs.EOF
REM 'add names seperated by comma to getNames
REM getNames = getNames + rs.Fields(0) & ","
REM 'move to next result before looping again
REM 'this is important
REM rs.MoveNext
REM 'continue loop
REM Loop
 
REM 'Close connection and release objects
REM objConn.Close
REM Set rs = Nothing
REM Set objConn = Nothing
 
REM 'Return Results via MsgBox
REM MsgBox getNames


sFilename = "C:\VB Script Training\test.accdb"  
Set objCN = CreateObject("ADODB.Connection") 
sConnection = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & sFilename & ";Persist Security Info = False;"
objCN.Open sConnection
sQuery="SELECT * FROM Employee" 
'Create recordset object
	Set objRs = CreateObject("ADODB.Recordset")
	objRs.Open sQuery, objCN, 2, 3
	'msgbox "Record count=" & objRS.RecordCount
	If Not objRS.EOF Then 
		arrRS = objRS.GetRows() 
	End If
	msgbox LBound(arrRS,2)
	msgbox UBound(arrRS,2)
	
	If IsArray(arrRS) Then
		For iRow=LBound(arrRS,2) to UBound(arrRS,2) ' for iRow= 0 to 
			Emp_id = arrRS(0,iRow)   'arrRS(0,0)
			Ename = arrRS(1,iRow)    
			Designation = arrRS(2,iRow)
			salary = arrRS(3,iRow)
			company = arrRS(4,iRow)
			msgbox Emp_id
			msgbox Ename 
			msgbox Designation
			msgbox salary
			msgbox company
		next
	end if
	Set objRS = Nothing
	Set objCn = Nothing
	
' insert data


REM set cmd = CreateObject("ADODB.Command")
REM 'set cmd.ActiveConnection = conn
REM cmd.commandText="Insert into Employee values (5,'Test5','Test Analyst','2000')"
REM cmd.execute
REM Set objRS = Nothing
REM Set objCn = Nothing
REM set cmd = Nothing

'insert data using

REM objRs.AddNew
REM objRS("ID").value = "6"
REM objRS("Emp_Name").value = "Test5"
REM objRS("Designation").value = "Test Analyst"
REM objRS("Salary").value = "5000"
REM objRS("Company").value = "SQS"
REM on error resume next
REM objRs.Update  
REM if err.number <> 0 then
REM msgbox err.description
REM end if
REM Set objRS = Nothing
REM Set objCn = Nothing