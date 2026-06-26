Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
appDir = scriptDir & "\pomodoro-app\dist\win-unpacked"

Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = appDir
WshShell.Run """" & appDir & "\番茄钟.exe""", 1, False
Set WshShell = Nothing
Set fso = Nothing
