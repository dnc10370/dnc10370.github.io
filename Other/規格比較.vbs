
StartURL ="http://mobileweb.tcc.net.tw/TCC/WebAP/mobile/index_special.html"
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
objIE.Width=1053         '設定IE物件寬度
objIE.Height=663         '設定IE物件高度
objIE.top=92
objIE.left=0
objIE.resizable=1        '設定IE物件大小是否可以被改動

