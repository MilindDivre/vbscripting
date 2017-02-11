'read data in file
Function readData(strFile)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFile, ForReading)
	Const ForReading = 1
	Do Until objFile.AtEndOfStream
		line=objFile.Readall
		msgbox LTrim(line)
	Loop
	objFile.Close
end function
call readData("C:\Users\divrem\Desktop\VB Script Training\test.qfl")

'Write Data

REM function writeData(strFile,strText)
REM set objFSO = CreateObject("Scripting.FileSystemObject")
REM set objFile=objFSO.createTextFile(strFile,false)
REM objFile.write(strText)

REM objFile.close
REM msgbox "done"
REM end function

'call writeData("c:\File1.txt","Hello1 World")

REM 'Append data


REM function appendData(strFile,strText)
REM set objFSO = createObject("Scripting.FileSystemObject")
REM on error resume next
REM set objFile = objFSO.OpenTextFile(strFile,8)

	REM if err.number <> 0  then
		REM msgbox "In valid file location"
	REM else
		REM objFile.write(vbCrlf & strText)
	REM end if
	

REM end function
REM call appendData("c:\File1.txt","Hello2 World")


