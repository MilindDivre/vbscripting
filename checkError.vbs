On Error Resume Next ' Turn on the error handling flag
   
x=1/0
Call CheckError()
sub checkError()
	Msgbox "In"
	if err.number > 0 then
		'Msgbox  "There is no error during runtime"
		Msgbox  err.description
	else
		Msgbox  "There is no error during runtime"
	end if
end sub

REM Err.Number Err.Description  

REM 5          Invalid procedure call or argument
REM 6          Overflow
REM 7          Out of Memory
REM 9          Subscript out of range
REM 10         This array is fixed or temporarily locked
REM 11         Division by zero
REM 13         Type mismatch
REM 14         Out of string space
REM 17         Can't perform requested operation
REM 28         Out of stack space
REM 35         Sub or function not defined
REM 48         Error in loading DLL
REM 51         Internal error
REM 91         Object variable not set
REM 92         For loop not initialized
REM 94         Invalid use of Null
REM 424        Object required