
'The script expects 3 arguments, version, build & scriptFile
If Wscript.Arguments.Count = 3 Then
	version = Wscript.Arguments(0)
	build = Wscript.Arguments(1)
	scriptFile = Wscript.Arguments(2)
Else
	WScript.Echo "Wrong no of arguments passed...."
	WScript.quit(1)
End If

'WScript.Echo "Master . vbs"
'Wscript.Echo version &" "&build&" "&scriptFile

'Set objWSHShell = CreateObject("Shell.Application")
Set WshShell = CreateObject("WScript.Shell")

'variables
globalFolderPath = "C:\automation\ApplicationSpecific\Scripts\Global\"
regressionFolderPath = "C:\automation\ApplicationSpecific\Scripts\Regression\"

Select case scriptFile
	case "SmokeQuicken.t"
		filePath = globalFolderPath&"SmokeQuicken.t"
	case "AccountList.t"
		filePath = regressionFolderPath&"""Account List\AccountList.t"""
	case "Business.t"
		filePath = regressionFolderPath&"Business\Business.t"
	case "BillManagement.t"
		filePath = regressionFolderPath&"""Bill Management\BillManagement.t"""
	case "ACE.t"
		filePath = regressionFolderPath&"ACE\ACE.t"
	case "Budget.t"
		filePath = regressionFolderPath&"Budget\Budget.t"
	case "C2R.t"
		filePath = regressionFolderPath&"Compare2Register\C2R.t"
	case "FileBackupRestore.t"
		filePath = regressionFolderPath&"""File Management\FileBackupRestore.t"""
	case "FileImportExport.t"
		filePath = regressionFolderPath&"""File Management\FileImportExport.t"""
	case "FileInputOutput.t"
		filePath = regressionFolderPath&"""File Management\FileInputOutput.t"""
	case "FileManagement.t"
		filePath = regressionFolderPath&"""File Management\FileManagement.t"""
	case "FileOperations.t"
		filePath = regressionFolderPath&"""File Management\FileOperations.t"""
	case "MenuNavigation.t"
		filePath = regressionFolderPath&"Generic\MenuNavigation.t"
	case "HomeTab.t"
		filePath = regressionFolderPath&"HomeTab\HomeTab.t"
	case "InvestingRegistersAndFormsPart1.t"
		filePath = regressionFolderPath&"""Investing Registers\InvestingRegistersAndFormsPart1.t"""
	case "InvestingRegistersAndFormsPart2.t"
		filePath = regressionFolderPath&"""Investing Registers\InvestingRegistersAndFormsPart2.t"""
	case "LoansPart1.t"
		filePath = regressionFolderPath&"Loans\LoansPart1.t"
	case "LoansPart2.t"
		filePath = regressionFolderPath&"Loans\LoansPart2.t"
	case "LoansPart3.t"
		filePath = regressionFolderPath&"Loans\LoansPart3.t"
	case "PropertyDebt_Part1.t"
		filePath = regressionFolderPath&"PropertyDebt\PropertyDebt_Part1.t"
	case "SpendingTab.t"
		filePath = regressionFolderPath&"SpendingTab\SpendingTab.t"
	case "RentalProperty.t"
		filePath = regressionFolderPath&"RPM\RentalProperty.t"
	case "SyncTest.t"
		filePath = regressionFolderPath&"QDSync\SyncTest.t"
	case "CurrencyList.t"
		filePath = regressionFolderPath&"""Currency List\CurrencyList.t"""
	case "DataSync.t"
		filePath = regressionFolderPath&"""Data Sync\DataSync.t"""
	case "Beacon_Testing.t"
		filePath = regressionFolderPath&"Generic\Beacon_Testing.t"
	case "DataConversion.t"
		filePath = regressionFolderPath&"Generic\DataConversion.t"
	case "HiddenAccount_Part1.t"
		filePath = regressionFolderPath&"""Hidden Account\HiddenAccount_Part1.t"""
	case "HiddenAccount_Part2.t"
		filePath = regressionFolderPath&"""Hidden Account\HiddenAccount_Part2.t"""
	case "memorized_payee.t"
		filePath = regressionFolderPath&"""Memorized Payee\memorized_payee.t"""
	case "PortfolioRebalancer.t"
		filePath = regressionFolderPath&"""Portfolio Rebalancer\PortfolioRebalancer.t"""
	case "Preferences.t"
		filePath = regressionFolderPath&"Preferences\Preferences.t"
	case "PropertyDebt_Part1.t"
		filePath = regressionFolderPath&"PropertyDebt\PropertyDebt_Part1.t"
	case "RegisterPart1.t"
		filePath = regressionFolderPath&"Register\RegisterPart1.t"
	case "RegisterPart2.t"
		filePath = regressionFolderPath&"Register\RegisterPart2.t"
	case "RegisterPart3.t"
		filePath = regressionFolderPath&"Register\RegisterPart3.t"
	case "RegisterPart4.t"
		filePath = regressionFolderPath&"Register\RegisterPart4.t"
	case "Registration.t"
		filePath = regressionFolderPath&"Registration\Registration.t"
	case "Renaming_Rules.t"
		filePath = regressionFolderPath&"""Renaming Rules\Renaming_Rules.t"""
	case "SavingGoals.t"
		filePath = regressionFolderPath&"SavingsGoal\SavingGoals.t"
	case "ScheduledTransactionsPart1.t"
		filePath = regressionFolderPath&"""Scheduled Transactions\ScheduledTransactionsPart1.t"""
	case "ScheduledTransactionsPart2.t"
		filePath = regressionFolderPath&"""Scheduled Transactions\ScheduledTransactionsPart2.t"""
	case else
		WScript.Echo "Invalid ScriptFile Name - "&scriptFile
		WScript.quit(1)
End Select

'uninstall Quicken
WshShell.Run "C:\quicken_build_job\UninstallQuicken.vbs",,true

'Install Quicken
'objWSHShell.ShellExecute "cscript.exe", Chr(34) & Chr(34) & "C:\quicken_build_job\BuildJob.vbs" & Chr(32) & version &" "&build, "", "runas", 1 
WshShell.Run "C:\quicken_build_job\BuildJob.vbs "& version &" "&build,,true

' Get automation code
'objWSHShell.ShellExecute "cscript.exe", Chr(34) & Chr(34) & "C:\quicken_build_job\Automation_Setup_job.vbs", "", "runas", 1 
WshShell.Run "C:\quicken_build_job\Automation_Setup_job.vbs "& version,,true

' Execute the silktest script
'objWSHShell.ShellExecute "cscript.exe", Chr(34) & Chr(34) & "C:\quicken_build_job\Silk_Execution_Controller.vbs" & Chr(32) & filePath, "", "runas", 1
WshShell.Run "C:\quicken_build_job\Silk_Execution_Controller.vbs " &filePath,,true

' copy results & logs from local archive folder in Z:
WshShell.Run "C:\quicken_build_job\Archive_Results_Logs.vbs",,true

WScript.quit
