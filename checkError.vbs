On Error Resume Next ' Turn on the error handling flag
   
x=1/0
'msgbox "ad"
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
' VBScript itself returns this error when it can't contact a remote machine. VBScript error numbers are all less than 10,000 decimal. Run-time errors are either less than 1,000 or between 5,000 and 5,100, while syntax errors are between 1,000 and 1,100. WMI and ADSI errors use larger numbers, generally 8-digit hex numbers.
' For scripts designed to run aga

' Const InvalidCall = 5
' Print “Global code start”
' Blah1
' Print “Global code end”
' Sub Blah1()
      ' On Error Resume Next
      ' Print “Blah1 Start”
      ' Blah2
      ' Print “Blah1 End”
' End Sub
' Sub Blah2()
      ' Print “Blah2 Start”       
      ' Err.Raise InvalidCall
      ' Print “Blah2 End”
' End Sub

' This prints out


' Global code start
' Blah1 Start
' Blah2 Start
' Blah1 End
' Global code end

' Hold on a minute — when the error happened, Blah1 had already turned ‘resume next’ mode on. The next statement after the error raise is Print “Blah2 End” but that statement never got executed. What’s going on?

' What’s going on is that the error mode is on a per-procedure basis, not a global basis. (If it were on a global basis, all kinds of bad things could happen — think about how you’d have to design a program to have consistent error handling in a world where that setting is global, and you’ll see why it’s per-procedure.) In this case, Blah2 gets an error. Blah2 is not in ‘resume next’ mode, so it aborts itself, records that there was an error situation, and returns to its caller. The caller sees the error, but the caller is in ‘resume next’ mode, so it resumes.

' In short, the propagation model for errors in VBScript is basically the same as traditional structured exception handling — the exception is thrown up the stack until someone catches it, or the program terminates. However, the error information that can be thrown