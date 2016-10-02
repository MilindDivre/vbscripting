REM Dim ToAddress
Dim FromAddress
Dim MessageSubject
Dim MyTime
Dim MessageBody
Dim MessageAttachment
Dim ol, ns, newMail
MyTime = Now

ToAddress = "milind.divre@sqs.com"
MessageSubject = "It works!."
MessageBody = "<!DOCTYPE html><html>"
 MessageBody = MessageBody &"<h1>Hello World<h1>"
 MessageBody = MessageBody &"</html>"
'MessageAttachment = "c:\File.txt"
'MessageAttachment1 = "c:\test.xls"
Set ol = CreateObject("Outlook.Application")
'Set ns = ol.getNamespace("MAPI")
Set newMail = ol.CreateItem(olMailItem)
newMail.Subject = MessageSubject
newMail.Body = MessageBody & vbCrLf & MyTime
newMail.RecipIents.Add(ToAddress)
'newMail.Attachments.Add(MessageAttachment)
'newMail.Attachments.Add(MessageAttachment1)
newMail.Send
Set objFolder    = Nothing
Set objNamespace = Nothing
Set objOutlook   = Nothing
msgbox "mail sent"

REM Dim olApp ''Outlook.Application
REM Dim olMapi ''Outlook.NameSpace
REM Dim olFolder ''Outlook.MAPIFolder
REM Dim olItems ''Outlook.Items

REM olFolderContacts = 10
REM Const olFolderInbox = 6

REM Set olApp = CreateObject("Outlook.Application")
REM Set olMapi = olApp.GetNamespace("MAPI")
REM Set olFolder = olMapi.GetDefaultFolder(olFolderInbox)
REM Set olItems = olFolder.Items

REM For i = 1 To olItems.Count
REM s = olItems(i)
REM msgbox s
REM Next 

REM MsgBox s
