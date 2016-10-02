Set QCConnection = CreateObject("TDApiOle80.TDConnection")
sQCUrl = "https://almtrafigura.saas.hp.com/qcbin"
QCConnection.InitConnectionEx  sQCUrl 		

QCConnection.Login "milind.divre", "Aditya786$"

QCConnection.Connect "TRADING_IT", "Titan"

If (QCConnection.LoggedIn <> True) Then
    MsgBox "QC User Authentication Failed"
    WScript.Quit
End If
msgbox "Connected toQC"