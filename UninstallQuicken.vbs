Set objShell = CreateObject( "WScript.Shell" )
appDataLocation=objShell.ExpandEnvironmentStrings("%APPDATA%")
appDataLocation = Replace(appDataLocation,"\Roaming","")


arg1 = appDataLocation&"\Local\Intuit"
arg2 = appDataLocation&"\Local\Intuit_Inc"
arg3 = appDataLocation&"\Roaming\Intuit"


objShell.Run "c:\quicken_build_job\Uninstall_Quicken.bat" &" "&arg1 &" "&arg2 &" "&arg3,,true



