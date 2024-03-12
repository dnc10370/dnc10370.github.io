
Set ws = CreateObject("Wscript.Shell")
ws.exec "C:\Program Files\Google\Chrome\Application\chrome.exe --app=https://scms.tcc.net.tw/sso/"
Wscript.Sleep 2000
For i= 0 to 1
Set WshShell = CreateObject("WScript.Shell")
WshShell.AppActivate "SCMS SSO" 
Wscript.Sleep 10
next
WshShell.SendKeys "{TAB}"
For i= 0 to 3
Wscript.Sleep 10
WshShell.SendKeys "H124196305"
Wscript.Sleep 10
WshShell.SendKeys "{TAB}"
Wscript.Sleep 10
WshShell.SendKeys "Aa123456"
Wscript.Sleep 10
WshShell.SendKeys "{TAB}"
Wscript.Sleep 10
WshShell.SendKeys "{TAB}"
Wscript.Sleep 10
WshShell.SendKeys "{TAB}"
Wscript.Sleep 10
WshShell.SendKeys "{TAB}"
Wscript.Sleep 10
next
WshShell.SendKeys "{TAB}"
Wscript.Sleep 10
WshShell.SendKeys "{TAB}"
Wscript.Sleep 10