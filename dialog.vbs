 REM a=inputBox("Enter number") 
REM msgbox a
REM a = MsgBox("Do you like blue color?",1,"Choose options")
REM msgbox a 
REM x=msgbox("Your Text Here" ,0, "Your Title Here")

REM Make sure that the " is included in the text and replace the Your Text Here and Your Title Here. But don't change anything elese!

REM Advanced users can change something else.
REM 0 =OK button only
REM 1 =OK and Cancel buttons
REM 2 =Abort, Retry, and Ignore buttons
REM 3 =Yes, No, and Cancel buttons
REM 4 =Yes and No buttons
REM 5 =Retry and Cancel buttons
REM 16 =Critical Message icon
REM 32 =Warning Query icon
REM 48 = Warning Message icon
REM 64 =Information Message icon
REM 0 = First button is default
REM 256 =Second button is default
REM 512 =Third button is default
REM 768 =Fourth button is default
REM 0 =Application modal (the current application will not work until the user responds to the message box)
REM 4096 =System modal (all applications wont work until the user responds to the

dialogArr =  array(0,1,2,3,4,16,32,48,64)
for each i in dialogArr
	msgbox "Your Text Here" ,i, "Your i="&i
next