on error resume next
StartURL ="https://scms.tcc.net.tw/ExtraApp/convergenceAddrQuery.do?module=EXTRA"
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
objIE.Width=700          '設定IE物件寬度
objIE.Height=400         '設定IE物件高度
objIE.top=252
objIE.left=280
objIE.resizable=1        '設定IE物件大小是否可以被改動
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop

set WshShell = CreateObject("WScript.Shell")

Wscript.Sleep 200
WshShell.SendKeys "{Tab}"
Wscript.Sleep 20
WshShell.SendKeys "{pgdn}"
Wscript.Sleep 20
WshShell.SendKeys "{up}"
Wscript.Sleep 20
WshShell.SendKeys "{up}"
Wscript.Sleep 20
WshShell.SendKeys "{up}"
Wscript.Sleep 20
WshShell.SendKeys "{up}"
Wscript.Sleep 100
WshShell.SendKeys "{Tab}"
Wscript.Sleep 200
WshShell.SendKeys "{pgdn}"


