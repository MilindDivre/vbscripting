Dim x,y
  x = 10
  y = 20
  z = fnadd(x,y)
  on error resume next
  a = fnmultiply(x,y)
  if err.number <> 0 then
	msgbox Err.Number & " Srce: Function does  " & Err.Source & " Desc: " &  Err.Description
  end if
  a=10
  b=0
 on error resume next
  d = a/b
  if err.number <> 0 then
	msgbox "Divede by zero err"
  end if
  msgbox "After error"
  
  Function fnadd(x,y)
      fnadd = x+y
  End Function