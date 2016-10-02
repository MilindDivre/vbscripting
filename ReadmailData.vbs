'how to read mail from outlook

Set objoutlook=CreateObject("Outlook.application")
Set objNamespace=objoutlook.getNameSpace("MAPI")
Set oFolder = objNamespace.GetDefaultFolder(6) 
Set allEmails = oFolder.Items 
For Each email In oFolder.Items 
	If email.Subject="Your Mail ID and Login account for SQS India is Created - Please Reply back for confirmation." then
	strbody= email.body 
	arrbody=split(strbody," ")
	msgbox Ubound(arrbody)
	msgbox arrbody(0)
		For i = 0 to Ubound(arrbody)
			If arrbody(i)="https://mail1.sqs.com/owa" then
				CreateObject("WScript.Shell").Run "https://mail1.sqs.com/owa"
			End If
		Next
End If		
Next


