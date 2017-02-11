Function readData(strFile)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFile, ForReading)
	Const ForReading = 1
	Do Until objFile.AtEndOfStream
		line=objFile.ReadLine
		'msgbox LTrim(line)
		'msgbox instr(1,LTrim(line),"function",1)
		if (instr(1,LTrim(line),"function",1)=1) then
			'msgbox "getting valid function name"
			msgbox LTrim(line)
			arrfunctionName=split(line,"function",-1,vbTextCompare)
			'msgbox "function name" &arrfunctionName(1)
			strfun=strClean(arrfunctionName(1))
			msgbox strfun
		end if
	Loop
	objFile.Close
end function

call readData("C:\Users\divrem\Desktop\VB Script Training\testfunctiondata.vbs")



Function strClean (strtoclean)
Dim objRegExp, outputStr
Set objRegExp = New Regexp

objRegExp.IgnoreCase = True
objRegExp.Global = True
objRegExp.Pattern = "(\(.*\))"
outputStr = objRegExp.Replace(strtoclean, "")
strClean = outputStr
End Function