Dim FSO, FSO2, logFolderPath, logFolder
logFolderPath = "c:\log"
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FSO2 = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
sourcePath = "c:\Latest\automation"
destinationPath = "c:\automation"

' Set the script engine to cscript
'WshShell.Run "cscript.exe //H:cscript"
'WScript.Echo "------------------------------------------------------------------"

'Terminate partner/runtime processes..
On Error Resume Next
WshShell.Run "taskkill /f /im partner.exe",,false
Wscript.Sleep 0.2*60*1000
WshShell.Run "taskkill /f /im runtime.exe",,false
Wscript.Sleep 0.2*60*1000
On Error GoTo 0

' Write into a log file
Set objFSO=CreateObject("Scripting.FileSystemObject")
writeLog = "c:\log\automation_setup_job.inf"
If Not objFSO.FolderExists(logFolderPath) Then
	Set logFolder = objFSO.CreateFolder(logFolderPath)
End If	
set objLogFile = objFSO.CreateTextFile(writeLog,True)


' verify SourceFolder exists or not..
If FSO.FolderExists(sourcePath) Then
	objLogFile.WriteLine "Automation Source folder Exists"
	'WScript.Echo "Automation Source folder Exists"
Else
	
	'WScript.Echo "Error~~~: Source Folder Does not Exists ->> " & sourcePath
	objLogFile.WriteLine "Error~~~: Source Folder Does not Exists ->> " & sourcePath
	objLogFile.Close
	WScript.quit(1)
End If

' Clean up.. if destinationFolder exists delete it
If FSO2.FolderExists(destinationPath) Then
	'WScript.Echo "Target Folder Exists, Deleting it for clean-up"
	'WScript.Echo "RD /s /q "&destinationPath
	objLogFile.WriteLine "Automation Folder Exists, Deleting it " & destinationPath
	WshShell.Run "c:\quicken_build_job\del.bat" 
	'FSO2.DeleteFolder (destinationPath)
End If

' copy the file
' VBS method CopyFolder throwing file not found exception, hence copying the folder 
' thru bat commands...
Set sourceFolder = FSO.GetFolder(sourcePath)
'FSO.CopyFolder sourcePath, destinationPath
WshShell.Run "c:\quicken_build_job\copy.bat " &sourcePath & " " &destinationPath,,true

' Verify whether the file copied successfully or not
'Wscript.Echo "Sleep for few Secs.."
Wscript.Sleep 0.8*60*1000
Set destinationFolder = FSO2.GetFolder(destinationPath)
'Wscript.Echo "Size of the source folder to copy " & sourceFolder.Size/1024/1024 &"MB"
'Wscript.Echo "Size of the destination folder to copy " & destinationFolder.Size/1024/1024 &"MB"
objLogFile.WriteLine "Size of the source folder to copy " & sourceFolder.Size/1024/1024 &"MB"
objLogFile.WriteLine "Size of the destinationFolder copied- " & destinationFolder.Size/1024/1024 &"MB"
'WScript.Echo destinationFolder.Size/1024/1024 &"MB"

if (CInt(sourceFolder.Size/1024/1024) = CInt(destinationFolder.Size/1024/1024)) Then
	'WScript.Echo "Automation Script folder Copied Successfully"
	objLogFile.WriteLine "Automation Script folder Copied Successfully to "& destinationPath
Else
	'WScript.Echo "Error~~~: Automation Folder Copy Failed ... :-("
	objLogFile.WriteLine "Error~~~: Automation Folder Copy Failed ... :-( but still continuing"
	'objLogFile.Close
	'WScript.quit(1)
End If

'WScript.Echo r
'Wscript.Echo "------------------------------------------------------------------"
objLogFile.Close
WScript.quit