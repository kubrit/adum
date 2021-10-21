command = "powershell.exe -executionpolicy bypass -WindowStyle Hidden -file \\example.domain\files\scripts\ADUM\adum.ps1"
set shell = CreateObject("WScript.Shell")
shell.Run command,0