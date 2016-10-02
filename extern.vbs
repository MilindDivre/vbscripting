Extern.Declare micHwnd, "FindWindow", "user32.dll", "FindWindowA", micString, micString
hwnd = Extern.FindWindow("Notepad", vbNullString)   
   
 If hwnd = 0 Then   
   MsgBox "Notepad window not found"   
  Else  
   MsgBox "Notepad window found"   
 end if   