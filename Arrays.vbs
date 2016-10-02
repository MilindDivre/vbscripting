' array can store numbers string as well
'######################################################################################################################################
REM arr = array("heloo","1","world","2","5.4",true)
REM for i = 0 to ubound(arr)
REM msgbox arr(i)
REM next
'######################################################################################################################################

'Dynamic arrays
dim weekDays()
Redim weekDays(3)
weekDays(0)="Monday"
weekDays(1)="Tuesday"
weekDays(2)="Wednesday"
weekDays(3)="Thursday"

for i = 0 to ubound(weekDays)
	 msgbox "In Array:" &weekDays(i)
next

redim weekdays(4)
weekDays(0)="Monday-Changed"
weekDays(1)="Tuesday-Changed"
weekDays(2)="Wednesday-Changed"
weekDays(3)="Thursday-Changed"
for i = 0 to ubound(weekDays)
	 msgbox "In Array:" &weekDays(i)
next
REM '######################################################################################################################################
REM 'Preserve array

Redim preserve weekDays(7)
weekDays(4)="Friday"
weekDays(5)="Saturday"
weekDays(6)="Sunday"
for each i in weekDays
	msgbox "In Array with Preserve"&i
next
REM '######################################################################################################################################
REM ' Join in Array
days = Join(weekDays,"*")
msgbox days
'######################################################################################################################################
'Split function
SplitWeekdaysArr = split(days,"*")
for each sday in SplitWeekdaysArr
msgbox "After array split:"&sday
next
'Multi Dimension array
REM dim arrMulti()
REM redim arrMulti(2,3)
REM redim preserve arrMulti(2,4)
REM msgbox ubound(arrMulti)
REM msgbox ubound(arrMulti,2)

REM dim multi(2,3) ' this will create a array of 3 rows and 4 columns
REM multi(0,0)="Emp_id"
REM multi(0,1)="Emp_name"
REM multi(0,2)="Emp_Sal"
REM multi(0,3)="Designation"

REM multi(1,0)="1"
REM multi(1,1)="test1"
REM multi(1,2)="10000"
REM multi(1,3)="Test Analyst"

REM multi(2,0)="2"
REM multi(2,1)="test2"
REM multi(2,2)="1000"
REM multi(2,3)="Analyst"

REM for irow=0 to 2
	REM for icol=0 to 3
		REM msgbox multi(irow,icol)
	REM next
REM next	
REM msgbox ubound(multi,2)