sFilename = "C:\VB Script Training\test.accdb"
ConnectionString = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=" & sFilename & ";"
'sql = "delete from Employee where ID = 1 " 
sql = "update Employee set Emp_Name = 'test updated' where ID = 3 " 
'sql = "Insert into Employee values (6,'Test5','Test Analyst','2000','SQS')" 
set cn = createobject("ADODB.Connection")
set cmd = createobject("ADODB.Command")
cn.open connectionString
cmd.ActiveConnection = cn
cmd.CommandText = sql
cmd.execute
cn.Close
msgbox "done!!"

