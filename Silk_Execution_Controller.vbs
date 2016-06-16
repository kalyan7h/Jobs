suiteFilePathToBeExecuted =  Wscript.Arguments(0)

' kick off the execution
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "C:\quicken_build_job\invokeSilkTestAndExecute.bat "&suiteFilePathToBeExecuted,,false

' wait for some time to start silktest and put somevalue inside it...
WScript.Sleep 1000*100


' wait for 4 hours max, if execution ends before look for the flag and teriminate silktest
for iCount = 1 to 16 step 1 
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile("C:\automation\stopper.txt", 1)
	strContents = objFile.ReadAll
	objFile.close
	
	bKillSilk = StrComp(Trim(Mid(strContents,1,4)),"Done",1)
	'Wscript.Echo bKillSilk
	'Wscript.Echo "["&Mid(strContents,1,4)&"]"
	
	if bKillSilk = 0 Then
		' kill silk process, execution is over
		'Wscript.Echo "matched...."
		Exit for
	End if
	
	' Wait for 15 mins 
	Wscript.Sleep 15*60*1000
	
	
Next

'Terminate partner/runtime & Agent processes, to send handle back to jenkins..
On Error Resume Next
WshShell.Run "taskkill /f /im partner.exe",,false
Wscript.Sleep 0.2*60*1000
WshShell.Run "taskkill /f /im runtime.exe",,false
Wscript.Sleep 0.2*60*1000
WshShell.Run "taskkill /f /im agent.exe",,false
On Error GoTo 0


	



