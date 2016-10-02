'create a XML File
'Set config file path
	REM sXmlpath = "C:\VB Script Training\Config.XML"
	REM 'Create file system object
	REM abc = "test"
	REM Set objFso = CreateObject("Scripting.FileSystemObject")
	REM set objTF=objFso.OpenTextFile (sXmlpath,2,True)
	REM 'Save all config settings as XML tags in file
	REM objTF.writeline "<Configuration>"
	REM objTF.Writeline "<Path>" & abc & "</Path>"
	REM objTF.Writeline "<Execute>" & abc & "</Execute>"
	REM objTF.WriteLine "<PageLoadTime>" & abc & "</PageLoadTime>"
	REM objTF.WriteLine "<EmailSend>" & abc & "</EmailSend>"
	REM objTF.WriteLine "<EmailTo>" & abc & "</EmailTo>"
	REM objTF.WriteLine "<Emailcc>" & abc & "</Emailcc>"
	REM objTF.WriteLine "<Emailbcc>" & abc & "</Emailbcc>"
	REM objTF.WriteLine "<EmailBody>" & abc & "</EmailBody>"
	REM objTF.WriteLine "<SendAttachment>" & abc & "</SendAttachment>"
	REM objTF.WriteLine "</Configuration>"
	REM 'Close file
	REM objTF.close
	REM 'Destroy objects
	REM Set objFso = Nothing
	REM Set objTF = Nothing
	
	'read XML file
set xmlDoc=CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
'change the xml file
xmlDoc.load("demo.xml")
'displays all the childNodes
REM for each x in xmlDoc.documentElement.childNodes
  REM msgbox(x.nodename &":"&x.text)
REM next

'for demo.xml
Set Root = xmlDoc.documentElement
Set NodeList = Root.getElementsByTagName("INTERFACE") 
For Each Elem In NodeList 
   SET port = Elem.getElementsByTagName("PORT")(0)
   SET ip = Elem.getElementsByTagName("IPADDRESS")(0)
   msgbox "Port " & port.text & " has IP address is " & ip.text
Next

REM Set colNodes=xmlDoc.selectNodes _
  REM ("//HARDWARE/COMPUTER[@os='Windows XP']")

REM For Each objNode in colNodes
  REM msgbox objNode.Text 
  REM msgbox objNode.Attributes.getNamedItem("department").text
REM Next
