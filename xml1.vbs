REM set objXML = CreateObject("microsoft.xmldom")
REM set objRoot =  objXML.createElement("Student")
REM objXML.appendChild objRoot
	REM set objRecord = objXML.createElement("Name")
	REM objRecord.text ="test1"
	REM objRoot.appendChild objRecord
	REM Set objIntro = objXML.createProcessingInstruction ("xml","version='1.0'")  
	REM objXML.insertBefore objIntro,objXML.childNodes(0)
 REM objXML.Save "C:\VB Script Training\test3.xml"
 REM msgbox "done!!"

 'Append data in XML file
 REM set objXML = CreateObject("Microsoft.xmldom")
 REM objXML.Async="False"
 REM objXML.load("test3.xml")
 REM set objRoot = objXML.documentElement
 REM set objStudent = objXML.createElement("Name")
 REM objStudent.text = "test2"
 REM objRoot.appendChild objStudent
 REM objXML.Save "C:\VB Script Training\test3.xml"
 REM msgbox "done!!"
 
 'modify the data from XML
 REM set objXML = createobject("Microsoft.xmldom")
 REM objXML.async ="False"
 REM objXML.load("test3.xml")
 REM set selNode = objXML.selectNodes("/Student/Name")
 REM for each node in selNode
	REM node.text ="Replaced_test"
 REM next
 REM objXML.Save "C:\VB Script Training\test3.xml"
 REM msgbox "done!"
 
 'delete data from xml
 REM set objXML = createobject("Microsoft.xmldom")
 REM objXML.async ="False"
 REM objXML.load("test3.xml")
 REM set selNode = objXML.selectNodes("/Student/Name")
 REM for each node in selNode
	REM objXML.documentElement.removechild(node)
 REM next
 REM objXML.Save "C:\VB Script Training\test3.xml"
 REM msgbox "done!"
 
 
 Set xmlDoc = _
  CreateObject("Microsoft.XMLDOM")  
  
Set objRoot = _
  xmlDoc.createElement("ITChecklist")  
xmlDoc.appendChild objRoot  

Set objRecord = _
  xmlDoc.createElement("ComputerAudit") 
objRoot.appendChild objRecord 
  
Set objName = _
  xmlDoc.createElement("ComputerName")  
objName.Text = "atl-ws-001"
objRecord.appendChild objName  

Set objDate = _
  xmlDoc.createElement("AuditDate")  
objDate.Text = Date  
objRecord.appendChild objDate  

Set objIntro = _
  xmlDoc.createProcessingInstruction _
  ("xml","version='1.0'")  
xmlDoc.insertBefore _
  objIntro,xmlDoc.childNodes(0)  

xmlDoc.Save "C:\VB Script Training\Audits.xml"  