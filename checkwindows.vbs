 REM dim objShell
        REM dim objShellWindows
        
        REM set objShell = CreateObject("shell.application")
        REM set objShellWindows = objShell.Windows

        REM if (not objShellWindows is nothing) then
            REM cnt = WScript.Echo objShellWindows.Count
			
        REM end if

        REM set objShellWindows = nothing
        REM set objShell = nothing
		
		
        dim objShell
        
        set objShell = CreateObject("shell.application")
        objShell.CascadeWindows
        set objShell = nothing
   