Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
appDir = fso.GetParentFolderName(WScript.ScriptFullName)
batPath = appDir & "\start.bat"
shell.Run "cmd.exe /c """ & batPath & """ --foreground", 0, False
