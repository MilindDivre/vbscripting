
  
    objOutlook = CType(CreateObject("Outlook.Application"), Outlook.Application)
    objEmail = objOutlook.CreateItem(Outlook.OlItemType.olMailItem)
    
    strb.Append("<table width='600px' align='center' border='0' cellpadding='0' cellspacing='0' style='border-top:5px solid white;'")
    strb.Append("<tr><td>S.No</td><td>AccountID</td><td>ChargeEntryControl</td><td>PaymentPostingControl</td></tr></table>")
    body = "Hi,"
    body = body & strb.ToString
    objEmail.htmlbody = body
    objEmail.display()