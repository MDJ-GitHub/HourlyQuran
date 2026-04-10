Set objShell = CreateObject("Wscript.Shell")

' Get script folder
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Run PowerShell hidden and set working directory
objShell.Run "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -Command ""Set-Location '" & scriptPath & "'; .\hourlyQuran.ps1""", 0, False