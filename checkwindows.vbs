  dim objShell
        dim objShellWindows
        
         set objShell = CreateObject("shell.application")
         set objShellWindows = objShell.Windows

         if (not objShellWindows is nothing) then
             cnt = objShellWindows.Count
			msgbox cnt
         end if

        set objShellWindows = nothing
         set objShell = nothing
		
		
        ' dim objShell
        
        ' set objShell = CreateObject("shell.application")
        ' objShell.CascadeWindows
        ' set objShell = nothing
   