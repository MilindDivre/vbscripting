Set regEx = New RegExp
regEx.Pattern="rain"
regEx.Global=true
regEx.Ignorecase=true

targetString = "The rain in Spain falls mainly in the plain Rain and raining " ' check for Ignore case
Set colMatch = regEx.Execute(targetString)
msgbox regEx.test(targetString) ' Check for invalid Patter
str= regEx.replace(targetString,"b") 'Check for replace
msgbox str
msgbox targetString ' No change in actual string

msgbox colMatch.Count
msgbox colMatch.Item(0)

for each match in colMatch
	msgbox match.FirstIndex
	msgbox match.length
	msgbox match.value
next

REM function submatchesex(strEmail)
REM set regex = new regexp
REM with regex
	REM .Pattern = "(\w+)@(\w+).(\w+)"
	REM .Global =true
	REM .Ignorecase =true
REM end	with
REM set exRegex = regex.execute(strEmail)
REM set match = exRegex.item(0)
REM msgbox "First name:" & match.submatches(0)
REM msgbox "Last name:" & match.submatches(1)
REM end function
REM call submatchesex("milind@gmail.com")
