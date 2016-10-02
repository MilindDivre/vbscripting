REM 'When calling sub you can use call or else don't use parenthesis 
REM function add(a,b)
REM c = a + b
REM 'add = c
REM msgbox "sum of "&a& "and  "&b& "="  &c
REM add = c
REM end function

REM d=add (byref7,byref2)
REM msgbox d
REM e=10
REM f=add(d,e)
REM msgbox f

'###############################################################################################

'ByRef passes the Reference , hence the value gets changed
'ByVal Passes the Value, hence the change is not noticed after the scope ends
'if the calling sub parameter is enclosed in parenthesis it will still pass by Val

 Sub TestSub(Byref MyParam) 
	 msgbox "In Sub" &MyParam
	 MyParam = 5
	 msgbox MyParam
 End Sub 

 Dim MyArg 
 MyArg = 123
 TestSub MyArg
 
 msgbox MyArg


'#######################################################################################################
'ByRef passes the Reference , hence the value gets changed
'ByVal Passes the Value, hence the change is not noticed after the scope ends
'if the calling sub parameter is enclosed in parenthesis it will still pass by Val

REM function TestSub(Byref MyParam) 
	REM msgbox MyParam
    REM MyParam = 5
	REM msgbox MyParam
REM End function

REM Dim MyArg1 
REM MyArg1 = 123
REM call TestSub (MyArg1)
REM msgbox MyArg1


'################################################################################################################