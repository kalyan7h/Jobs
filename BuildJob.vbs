
Dim FSO, FSO2, logFolderPath, logFolder
logFolderPath = "c:\log"
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FSO2 = CreateObject("Scripting.FileSystemObject")

' Write into a log file
Set objFSO=CreateObject("Scripting.FileSystemObject")
writeLog = "c:\log\buildjob.inf"
If Not objFSO.FolderExists(logFolderPath) Then
	Set logFolder = objFSO.CreateFolder(logFolderPath)
End If	
set objLogFile = objFSO.CreateTextFile(writeLog,True)

If WScript.Arguments.Count <> 2 Then
	objLogFile.WriteLine "Error~~~: BuildJob.vbs Expecting 2 arguments, Actual Passed " &WScript.Arguments.Count
	objLogFile.Close
	WScript.quit(1)
End If
	
version = Wscript.Arguments(0)
build=Wscript.Arguments(1)
Set WshShell = CreateObject("WScript.Shell")

' see qw.exe running in the process, kill the process
On Error Resume Next
WshShell.Run "taskkill /f /im qw.exe",,false
On Error GoTo 0

' Set the script engine to cscript
'WshShell.Run "cscript.exe //H:cscript"
'Wscript.Echo "------------------------------------------------------------------"
'Wscript.Echo version
'Wscript.Echo build

sourcePath = "z:\QuickenBuilds\QW"&version&"\qknty"&version&"baseinstaller-"&build&"-distribution\RPM"
'Wscript.Echo sourcePath
'sourcePath = "C:\QuickenBuilds\25.1.4.13\RPM"
destinationPath = "c:\X"



' verify SourceFolder exists or not..
If FSO.FolderExists(sourcePath) Then
	objLogFile.WriteLine "Source folder Exists"
	'WScript.Echo "Source Folder Exists"
Else
	
	'WScript.Echo "Source Folder Does not Exists"
	objLogFile.WriteLine "Error~~~: Source Folder Does not Exists ->> " & sourcePath
	objLogFile.Close
	WScript.quit(1)
End If

' Clean up.. if destinationFolder exists delete it
If FSO2.FolderExists(destinationPath) Then
	'WScript.Echo "Target Folder Exists, Deleting it for clean-up"
	objLogFile.WriteLine "Folder Exists, Deleting it" & destinationPath
	FSO2.DeleteFolder (destinationPath)
End If

' copy the file
Set sourceFolder = FSO.GetFolder(sourcePath)
FSO.CopyFolder sourcePath, destinationPath

' Verify whether the file copied successfully or not
Set destinationFolder = FSO2.GetFolder(destinationPath)
objLogFile.WriteLine "Size of the source folder to copy " & destinationFolder.Size/1024/1024 &"MB"
'WScript.Echo destinationFolder.Size/1024/1024 &"MB"

if (sourceFolder.Size = destinationFolder.Size) Then
	'WScript.Echo "Build Copied Successfully"
	objLogFile.WriteLine "Build Copied Successfully to "& destinationPath
Else
	'WScript.Echo "Copy Failed ... :-("
	objLogFile.WriteLine "Error~~~: Copy Failed ... :-("
	objLogFile.Close
	WScript.quit(1)
End If

' Copy successful, Instal quicken in the silent mode
'WScript.Echo "Quicken Installation Started.."
objLogFile.WriteLine "Quicken Installation Started.."
rpmInstallPath = destinationPath & "\Disk1\setup.exe /s"
WshShell.Run rpmInstallPath,,false

'wait for some 1> mins, hoping installation wont take more than 1 mins..
Wscript.Sleep 1.2*60*1000
'WScript.Echo "Checking whether qw process running or not..."

' see qw.exe running in the process even after installation, kill the process
On Error Resume Next
WshShell.Run "taskkill /f /im qw.exe",,false
objLogFile.WriteLine "Quicken Installation completed.."
On Error GoTo 0

'Wscript.Echo "------------------------------------------------------------------"
objLogFile.WriteLine "Quicken Installation completed"
objLogFile.Close
WScript.quit

