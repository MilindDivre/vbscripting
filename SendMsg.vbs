'Declaring the variables
Dim outobj,mailobj

Set outobj = CreateObject("Outlook.Application")
Set mailobj = outobj.CreateItem(0)


With mailobj
     .To = "+917775869775@ideacellular.net"
     .Subject = "Testmail"
     .Body = "Testmail"
     .Send
End With

'Clear the memory
Set outobj = Nothing
Set mailobj = Nothing