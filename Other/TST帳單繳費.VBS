
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.RegWrite"HKCU\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"
StartURL ="https://xws.tstartel.com/xws/initialBillQueryAndPay.do?chid=A"
Set Shell = WScript.CreateObject("WScript.Shell")
Set IE = WScript.CreateObject("InternetExplorer.Application", "IE_")
IE.Visible = 1
Shell.AppActivate IE

IE.Navigate StartURL  '設定IE物件預設指向的頁面 
IE.FullScreen=0       '全屏化IE物件
IE.menubar=0          '不顯示IE物件選單欄
IE.AddressBar=0       '不顯示IE物件位址列
IE.ToolBar=0          '不顯示IE物件工具欄
IE.StatusBar=0        '不顯示IE物件狀態列
IE.Width=1030         '設定IE物件寬度
IE.Height=653         '設定IE物件高度
IE.top=92
IE.left=0
IE.resizable=1        '設定IE物件大小是否可以被改動

