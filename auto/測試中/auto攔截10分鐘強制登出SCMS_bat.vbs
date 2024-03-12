


Set WshShell =WScript.CreateObject("WScript.Shell")

WshShell.Run"Taskkill /fi ""windowTitle eq SCMS SSO"" /T",0
