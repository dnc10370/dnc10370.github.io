Dim username,password,IE  ' 定義變量
'---------------------------------------------------------------對應網站
StartURL ="https://scms.tcc.net.tw/newetu/mOpQuery/OpQuery/queryRechargeTrans"

on error resume next
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE_" )
objIE.Visible = 1
objShell.AppActivate objIE
objIE.Navigate StartURL  '設定IE物件預設指向的頁面 
objIE.FullScreen=0      '全屏化IE物件
objIE.menubar=0         '不顯示IE物件選單欄
objIE.AddressBar=0       '不顯示IE物件位址列
objIE.ToolBar=0        '不顯示IE物件工具欄
objIE.StatusBar=1       '不顯示IE物件狀態列
objIE.Width=850         '設定IE物件寬度
objIE.Height=450         '設定IE物件高度
objIE.top=152         
objIE.left=130
objIE.resizable=1       '設定IE物件大小是否可以被改動

Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop
set WshShell = CreateObject("WScript.Shell")
Wscript.Sleep 850
objIE.document.getElementByid("search").click  
Wscript.Sleep 1100
objIE.document.getElementByid("search").click  


set WshShell = CreateObject("WScript.Shell")
Wscript.Sleep 850
WshShell.SendKeys "^p"
Wscript.Sleep 1050
WshShell.SendKeys "{Enter}"
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop
Wscript.Sleep 2000
kill4399 = "https://scms.tcc.net.tw/newetu/mOpQuery/OpQuery/queryRechargeTrans"
set sh = CreateObject("Shell.Application")
set wnds = sh.windows()
if Weekday(date)=0 Or Weekday(date)=0 then
end if
for each wnd in wnds
if InStr(1, wnd.LocationURL, kill4399, 1) then wnd.Quit()
Wscript.Sleep 50
next