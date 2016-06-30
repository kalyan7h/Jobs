Set objFSO = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

sYear = Year(Now)
sMonth = MonthName(Month(Now),False)
sDate = Day(Now)


currentTime = Now &"_Execution"
'Wscript.Echo currentTime
currentTime=Replace (currentTime,"/","-")
currentTime=Replace (currentTime," ","#")
currentTime=Replace (currentTime,":","_")

If Not objFSO.FolderExists("z:\Results_Archive") Then
	Set archiveFolder = objFSO.CreateFolder("z:\Results_Archive")
End If	

If Not objFSO.FolderExists("z:\Results_Archive\"&sYear) Then
	 Set archiveFolder = objFSO.CreateFolder("z:\Results_Archive\"&sYear)
End If

If Not objFSO.FolderExists("z:\Results_Archive\"&sYear&"\"&sMonth) Then
	 Set archiveFolder = objFSO.CreateFolder("z:\Results_Archive\"&sYear&"\"&sMonth)
End If

If Not objFSO.FolderExists("z:\Results_Archive\"&sYear&"\"&sMonth&"\"&sDate) Then
	 Set archiveFolder = objFSO.CreateFolder("z:\Results_Archive\"&sYear&"\"&sMonth&"\"&sDate)
End If

filePath = "z:\Results_Archive\"&sYear&"\"&sMonth&"\"&sDate&"\"&currentTime

If Not objFSO.FolderExists(filePath) Then
	 Set archiveFolder = objFSO.CreateFolder(filePath)
	 Wscript.Sleep 0.2*60*1000
End If

'copy results file to the shared location
WshShell.Run "xcopy c:\Results "&filepath &" /e /y"
Wscript.Sleep 0.2*60*1000

'copy logs
WshShell.Run "xcopy c:\log "&filepath &" /e /y"
Wscript.Sleep 0.2*60*1000

' clean-up - clear the local result folder 
objFSO.DeleteFile("c:\Results\*"), TRUE