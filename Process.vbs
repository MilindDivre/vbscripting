strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery _
 ("SELECT * FROM Win32_Process")
For Each objProcess in colProcessList
 msgbox objProcess.Name
Next
msgbox "done"