
'Const declaration
REM Const TIMEOUT = 54
REM Const MY_STRING_CONSTANT = "Hello World"
REM msgbox TIMEOUT
REM msgbox MY_STRING_CONSTANT
Const CutoffDate = #6-1-97#
MsgBox CutoffDate
'######################################################################################################################################
' Error while changing the value of const
REM TIMEOUT = 68
REM MY_STRING_CONSTANT = "Bye world"
REM msgbox TIMEOUT
REM msgbox MY_STRING_CONSTANT
'######################################################################################################################################
' Below code explains that we cannot assign  const a value that is return function
REM sVar = "Hello"
REM const sConst = len(sVar)
'######################################################################################################################################
REM Const START_TIME = 53 + 1
REM X = 10
REM Const START_TIME = 53 + X
 