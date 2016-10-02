Set objWord = CreateObject("Word.Application")
Set colTasks = objWord.Tasks
for each i in colTasks
	if (instr(i.name,"Outlook")) then
	msgbox "in if"
		i.close
	end if
next
set objWord =  nothing