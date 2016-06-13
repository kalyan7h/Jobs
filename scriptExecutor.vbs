Set objWSHShell = CreateObject("Shell.Application")
If Wscript.Arguments.Count = 2 Then
	version = Wscript.Arguments(0)
	build = Wscript.Arguments(1)
	objWSHShell.ShellExecute "cscript.exe", Chr(34) & Chr(34) & "C:\quicken_build_job\BuildJob.vbs" & Chr(32) & version &" "&build, "", "runas", 1 
Else
	objWSHShell.ShellExecute "cscript.exe", Chr(34) & Chr(34) & "C:\quicken_build_job\Automation_Setup_job.vbs", "", "runas", 1 
End If
