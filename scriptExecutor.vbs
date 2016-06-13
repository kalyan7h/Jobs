vbsFilePath = Wscript.Arguments(0)
Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cscript", vbsFilePath, "", "runas", 1
