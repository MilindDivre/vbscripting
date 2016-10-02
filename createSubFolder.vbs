with CreateObject("Shell.Application")
  set oFolder = .NameSpace("C:\")
  if (not oFolder is nothing) then oFolder.NewFolder("a\b\c\d")
end with