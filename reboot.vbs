strComputer = "." 

SET objWMIService = GETOBJECT("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\" & _
			strComputer & "\root\cimv2")

SET colOS = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
'Msgbox "Rebooting Machine...."
FOR EACH objOS in colOS
	objOS.Reboot()
NEXT