' create a folder
'if the folder exists it does not overrides
'if you provide path as mytest1\mytest2\mytest3 wont create 3 folders
REM function createFolder(strPath)
REM set fso = createObject("Scripting.FileSystemObject")
REM set fldr = fso.createFolder(strPath)
REM msgbox "folder created " & fldr.name
REM end function
REM 
REM for i = 0 to 10
REM foldername = i
REM call createFolder(foldername)
REM next

'Drive example
Function DriveDetails(strDriveName)
	Set oFS = CreateObject("Scripting.FileSystemObject")
	Set drive = oFS.GetDrive(oFS.GetDriveName(strDriveName))
	s = "Drive " & UCase(strDriveName) & " - " 
	s = s & "Free Space: " & FormatNumber(drive.FreeSpace/1024, 0) 
	s = s & " Kbytes"
	s = s & "FileSystem name:" &drive.FileSystem &vblf
	s = s & "AvailableSpace:" &drive.AvailableSpace &vblf
	s = s & "DriveLetter:" &drive.DriveLetter &vblf
	s = s & "DriveType:" &drive.DriveType &vblf
	s = s & "IsReady:" &drive.IsReady &vblf
	s = s & "Path:" &drive.Path &vblf
	s = s & "RootFolder:" &drive.RootFolder &vblf
	s = s & "SerialNumber:" &drive.SerialNumber &vblf
	s = s & "ShareName:" &drive.ShareName &vblf
	s = s & "TotalSize:" &drive.TotalSize &vblf
	s = s & "VolumeName:" &drive.VolumeName &vblf

	msgbox s
end function
call DriveDetails("c:\")

'Drives example
REM function getDrivesInfo()
REM set fso =  createObject("Scripting.FileSystemObject")
REM set fd = fso.drives
REM for each drive in fd
	REM infoDrives = "Drive Letter :" &drive.DriveLetter &vblf
	REM infoDrives = infoDrives & "Is Ready:" & drive.isReady
	REM msgbox infoDrives
REM next
REM end function
'call getDrivesInfo()

' Copy,Move, delete file
' file gets overridden by source if destination already exists
REM function scopyfile(sourcePath, destinationPath)
REM set fso =  createObject("Scripting.FileSystemObject")
REM fso.movefile "c:\mytest\test.txt",destinationPath
REM 'fso.copyfile "c:\mytest1\test.txt",destinationPath
REM 'fso.copyfile sourcePath,destinationPath 
REM 'fso.deletefile sourcePath
REM end function
REM call scopyfile("c:\mytest\test.txt","c:\mytest1\")



'File Attributes
' function getFileAttributes(strFilename)
' Dim fso, f
' Set fso = CreateObject("Scripting.FileSystemObject")
' Set f = fso.GetFile(strFilename)
' msgbox "Line 1: "& f.DateCreated &vblf
' msgbox "Line 2: "& f.Attributes &vblf
' msgbox "Line 3: "& f.DateLastAccessed &vblf
' msgbox "Line 4: "& f.DateLastModified &vblf
' msgbox "Line 5: "& f.Drive &vblf
' msgbox "Line 6: "& f.Name  &vblf
' msgbox "Line 7: "& f.ParentFolder &vblf 
' msgbox "Line 8: "& f.Path  &vblf
' msgbox "Line 9: "& f.ShortName  &vblf
' msgbox "Line 10: "& f.ShortPath 
' msgbox "Line 11: "& f.Size  
' msgbox "Line 12: "& f.Type 
' end function 
' call getFileAttributes("c:\mytest\test.txt")

REM Set objFSO = CreateObject("Scripting.FileSystemObject")
REM objStartFolder = "C:\VB Script Training"

REM Set objFolder = objFSO.GetFolder(objStartFolder)

REM Set colFiles = objFolder.Files
REM For Each objFile in colFiles
    REM Wscript.Echo objFile.Name
	
REM Next