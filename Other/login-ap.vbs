on error resume next
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.RegWrite"HKCU\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"
StartURL ="\\DNC10370-PC3\WinHotKey\VBS\SCMS常用連結.html"
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE_")
objIE.Visible = 1
objShell.AppActivate objIE

objIE.Navigate StartURL  '設定IE物件預設指向的頁面 
objIE.FullScreen=0       '全屏化IE物件
objIE.menubar=0          '不顯示IE物件選單欄
objIE.AddressBar=0       '不顯示IE物件位址列
objIE.ToolBar=0          '不顯示IE物件工具欄
objIE.StatusBar=0        '不顯示IE物件狀態列
objIE.Width=445         '設定IE物件寬度
objIE.Height=515         '設定IE物件高度
objIE.top=120
objIE.left=285

objIE.resizable=0        '設定IE物件大小是否可以被改動
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop
objIE.document.getElementByid("login-ap").click  

Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop

Wscript.Quit