REM set dict = CreateObject("Scripting.Dictionary")
REM dict.add "Italy","Rome"
REM dict.add "Germany","Berlin"
REM dict.add "India","Delhi"
REM dict.add "Britain","London"

REM msgbox dict.item("Germany")
'add the same key
'dict.add "India","Delhi123"
REM msgbox  "Count of items in Dictionary:" &dict.count
REM msgbox "Check if item exists:" &dict.exists("India")
'Dictionary object makes a binary comparison when it tries to find a match between the specified key and the stored keys.
 set dict1 = CreateObject("Scripting.Dictionary")

 'dict1.CompareMode = vbTextCompare 'Attempting to set this property for a nonempty dictionary results in an error.
 dict1.add "Italy","Rome"
 dict1.add "Germany","Berlin"
 dict1.add "India","Delhi"
 dict1.add "Britain","London"

 'msgbox "Check if item exists:" &dict1.exists("india")

 'msgbox dict1.item("asdah") ' key gets added to dictionary with null value
 'Enumerating dictionary items
 for each item in dict1
	 msgbox item & "->" &dict1.item(item)
next
REM 'keys
REM msgbox "Second key " &dict.keys()(1)
REM 'dict.remove("Italy")
REM 'dict1.removeall
REM dict.key("Germany")="Holland"
REM msgbox "Second key " &dict.keys()(1)



REM 'Function
REM function testRepaeat(strWord)
REM length = len(strWord)
REM Dim dict
REM Set dict = CreateObject ("Scripting.Dictionary")
REM for i = 1 to length
  REM alp=mid(strWord,i,1)
  REM 'msgbox alp
  REM abc=dict.Exists(alp)
  REM 'msgbox abc
  REM if abc then
    REM msgbox alp &" is repeated"
    REM else
    REM dict.Add alp, i
  REM end if
REM next
REM end function
REM 'call testRepaeat("hello how are you")


REM call called()
REM function callin()
REM Set cars = CreateObject("Scripting.Dictionary") 
REM cars.Add "a", "Alvis" 
REM cars.Add "b", "Buick" 
REM cars.Add "c", "Cadillac"
 REM set callin = cars
REM end function

REM function called()
REM set objCallin = callin()
REM msgbox "The value corresponding to the key 'b' is " & objCallin.Item("b") 

REM msgbox objCallin.count
REM msgbox objCallin.exists("a")
REM end function
