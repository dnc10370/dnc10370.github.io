on error resume next
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.Run"taskkill /f /im ielowutil.exe",0
kill4399 = "https://scms.tcc.net.tw/openchannel/checkExtLogin.do"
kill4400 = "https://scms.tcc.net.tw/sso/index.jsp"
set sh = CreateObject("Shell.Application")
set wnds = sh.windows()
if Weekday(date)=0 Or Weekday(date)=0 then
end if
For i= 0 to 10
for each wnd in wnds
if InStr(1, wnd.LocationURL, kill4399, 1) then wnd.Quit()
if InStr(1, wnd.LocationURL, kill4400, 1) then wnd.Quit()
next
Wscript.Sleep 50
next
Set WshShell =WScript.CreateObject("WScript.Shell")

WshShell.Run"Taskkill /fi ""windowTitle eq SCMS SSO - Internet Explorer"" /T",0

set objword = CreateObject("word.Application")
Set colTasks = objWord.Tasks

For i= 0 to 10
for each colTask in colTasks
if instr(colTask.name,"網頁訊息")<>0 then
colTask.close
end if
next
Wscript.Sleep 200
for each colTask in colTasks
if instr(colTask.name,"SCMS SSO - Internet Explorer")<>0 then
colTask.close
end if
next
Wscript.Sleep 50
for each colTask in colTasks
if instr(colTask.name,"系統訊息")<>0 then
colTask.close
end if
next
Wscript.Sleep 50
next


kill4402 = "https://scms.tcc.net.tw/sso/auth/certification.do?webAction="
set sh = CreateObject("Shell.Application")
set wnds = sh.windows()
if Weekday(date)=0 Or Weekday(date)=0 then
end if
for each wnd in wnds
if InStr(1, wnd.LocationURL, kill4402, 1) then wnd.Quit()
next


Start0URL ="https://scms.tcc.net.tw/sso/cca/index.jsp"
Set IE =CreateObject("internetexplorer.application")
IE.Visible =0         '設定是否可見
IE.Navigate Start0URL  '設定IE物件預設指向的頁面 
IE.FullScreen=0       '全屏化IE物件
IE.menubar=0          '不顯示IE物件選單欄
IE.AddressBar=0       '不顯示IE物件位址列
IE.ToolBar=0          '不顯示IE物件工具欄
IE.StatusBar=0        '不顯示IE物件狀態列
IE.Width=0        '設定IE物件寬度
IE.Height=0         '設定IE物件高度
IE.resizable=0        '設定IE物件大小是否可以被改動


IE.Quit




Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.Run"\\DNC10370-PC3\WinHotKey\VBS\auto\countdown.bat",0
