Dim FSO, FSO2
Set FSO = CreateObject("Scripting.FileSystemObject")
Set FSO2 = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")
sourcePath = "c:\Latest\automation"
destinationPath = "c:\automation"

' Write into a log file
Set objFSO=CreateObject("Scripting.FileSystemObject")
writeLog = "c:\log\automation_setup_job.inf"
set objLogFile = objFSO.CreateTextFile(writeLog,True)

' verify SourceFolder exists or not..
If FSO.FolderExists(sourcePath) Then
	objLogFile.WriteLine "Automation Source folder Exists"
	WScript.Echo "Source Folder Exists"
Else
	
	WScript.Echo "Source Folder Does not Exists"
	objLogFile.WriteLine "Error***: Source Folder Does not Exists ->> " & sourcePath
	objLogFile.Close
	WScript.quit
End If

' Clean up.. if destinationFolder exists delete it
If FSO2.FolderExists(destinationPath) Then
	WScript.Echo "Target Folder Exists, Deleting it for clean-up"
	WScript.Echo "RD /s /q "&destinationPath
	objLogFile.WriteLine "Automation Folder Exists, Deleting it" & destinationPath
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
Wscript.Echo "kalyan"
Wscript.Sleep 0.1*60*1000
Set destinationFolder = FSO2.GetFolder(destinationPath)
Wscript.Echo destinationFolder.Size
objLogFile.WriteLine "Size of the source folder to copy " & sourceFolder.Size/1024/1024 &"MB"
objLogFile.WriteLine "Size of the destinationFolder copied- " & destinationFolder.Size/1024/1024 &"MB"
'WScript.Echo destinationFolder.Size/1024/1024 &"MB"

if (CInt(sourceFolder.Size/1024/1024) = CInt(destinationFolder.Size/1024/1024)) Then
	WScript.Echo "Build Copied Successfully"
	objLogFile.WriteLine "Build Copied Successfully to "& destinationPath
Else
	WScript.Echo "Copy Failed ... :-("
	objLogFile.WriteLine "Error***: Copy Failed ... :-("
	objLogFile.Close
	WScript.quit
End If

'WScript.Echo r
objLogFile.Close