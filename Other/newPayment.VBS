on error resume next
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.RegWrite"HKCU\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"
StartURL ="https://scms.tcc.net.tw/newPayment/unfinishChargeProc.do?Init=Yes"
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
objIE.Width=1030         '設定IE物件寬度
objIE.Height=653         '設定IE物件高度
objIE.top=92
objIE.left=0
objIE.resizable=1        '設定IE物件大小是否可以被改動

Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop

set WshShell = CreateObject("WScript.Shell")

Wscript.Sleep 50

WshShell.SendKeys "{F6}"

WshShell.SendKeys "{TAB}"

WshShell.SendKeys "{TAB}"

WshShell.SendKeys "{TAB}"

WshShell.SendKeys "{DOWN}"

WshShell.SendKeys "{TAB}"

Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop

Wscript.Quit