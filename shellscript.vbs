Set shell = CreateObject("WScript.Shell")
shell.Run Chr(34) & "C:\AutoScriptRunner\script.bat" & Chr(34),0
shell.CurrentDirectory = "C:\AutoScriptRunner"