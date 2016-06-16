
'The script expects 3 arguments, version, build & scriptFile
If Wscript.Arguments.Count = 3 Then
	version = Wscript.Arguments(0)
	build = Wscript.Arguments(1)
	scriptFile = Wscript.Arguments(2)
Else
	WScript.Echo "Wrong no of arguments passed...."
	WScript.quit(1)
End If

WScript.Echo "Master . vbs"
Wscript.Echo version &" "&build&" "&scriptFile

'Set objWSHShell = CreateObject("Shell.Application")
Set WshShell = CreateObject("WScript.Shell")

'variables
globalFolderPath = "C:\automation\ApplicationSpecific\Scripts\Global\"
regressionFolderPath = "C:\Latest\automation\ApplicationSpecific\Scripts\Regression\"

Select case scriptFile
	case "SmokeQuicken.t"
		filePath = globalFolderPath&"SmokeQuicken.t"
	case "AccountList.t"
		filePath = regressionFolderPath&"Account List\AccountList.t"
	case "Business.t"
		filePath = regressionFolderPath&"Business\Business.t"
	case else
		WScript.Echo "Invalid ScriptFile Name - "&scriptFile
		WScript.quit(1)
End Select


'Install Quicken
'objWSHShell.ShellExecute "cscript.exe", Chr(34) & Chr(34) & "C:\quicken_build_job\BuildJob.vbs" & Chr(32) & version &" "&build, "", "runas", 1 
WshShell.Run "C:\quicken_build_job\BuildJob.vbs "& version &" "&build,,true

WScript.Echo "check Quicken Installation completed or not...."

' Get automation code
'objWSHShell.ShellExecute "cscript.exe", Chr(34) & Chr(34) & "C:\quicken_build_job\Automation_Setup_job.vbs", "", "runas", 1 
WshShell.Run "C:\quicken_build_job\Automation_Setup_job.vbs",,true

WScript.Echo "check Automation setup completed or not...."

' Execute the silktest script
'objWSHShell.ShellExecute "cscript.exe", Chr(34) & Chr(34) & "C:\quicken_build_job\Silk_Execution_Controller.vbs" & Chr(32) & filePath, "", "runas", 1
WshShell.Run "C:\quicken_build_job\Silk_Execution_Controller.vbs " &filePath,,true

WScript.Echo "is Exec completed??"

WScript.quit
