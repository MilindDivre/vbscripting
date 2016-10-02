
set dict = createObject("Scripting.Dictionary")
dict.add "Alfresco","http://alfrescocontint:8080/alfresco/faces/jsp/dashboards/container.jsp"
dict.add "Confluence","https://confluence/login.action"
dict.add "ALM","https://almtrafigura.saas.hp.com/qcbin/start_a.jsp"
dict.add "Starling","http://starlings/"
dict.add "Splunk","http://splunk-stag/en-GB/app/"
dict.add "Titan","http://titan-clickonce/"
Const ADMINISTRATIVE_TOOLS = 6 
Set objShell = CreateObject("Shell.Application") 
Set objFolder = objShell.Namespace(ADMINISTRATIVE_TOOLS)  
Set objFolderItem = objFolder.Self      
Set objShell = WScript.CreateObject("WScript.Shell") 
strDesktopFld = objFolderItem.Path 
for each item in dict
	Set objURLShortcut = objShell.CreateShortcut(strDesktopFld & "\"& item &".url") 
	objURLShortcut.TargetPath = dict.item(item) 
	objURLShortcut.Save 
next
msgbox "Favorites Added Close IE And Relaunch Again!!"