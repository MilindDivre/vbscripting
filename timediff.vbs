StartTime = Timer()
msgbox StartTime
EndTime = Timer()
msgbox EndTime
msgbox("Seconds to 2 decimal places: " & FormatNumber(EndTime - StartTime, 2))