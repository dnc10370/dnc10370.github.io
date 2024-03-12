Set wshobj=WScript.CreateObject("WScript.Shell")
code="新北市鶯歌區"
wshobj.Run "cmd.exe /c echo " & code & "| clip.exe", vbHide
WScript.Sleep 300
wshobj.AppActivate app
wshobj.SendKeys "^v"
set WshShell = CreateObject("WScript.Shell")
WScript.Sleep 50
WshShell.SendKeys "{TAB}"
WScript.Sleep 50
WshShell.SendKeys "{TAB}"
WScript.Sleep 50
WshShell.SendKeys "{Enter}"
WScript.Sleep 100
WshShell.SendKeys "{ESC}"
wscript.Quit