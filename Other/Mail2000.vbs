on error resume next
Dim username,password,IE  ' 定義變量
username = "dnc10370_1@fc.com"
password = "Zz09350935"
StartURL ="http://webmail.fc.com/cgi-bin/login?index=1"
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE_" )
objIE.Visible = 1
objShell.AppActivate objIE
objIE.Navigate StartURL  '設定IE物件預設指向的頁面 
objIE.FullScreen=0      '全屏化IE物件
objIE.menubar=0        '不顯示IE物件選單欄
objIE.AddressBar=0      '不顯示IE物件位址列
objIE.ToolBar=0         '不顯示IE物件工具欄
objIE.StatusBar=0       '不顯示IE物件狀態列
objIE.Width=1035         '設定IE物件寬度
objIE.Height=740         '設定IE物件高度
objIE.top=0         
objIE.left=0
objIE.resizable=1       '設定IE物件大小是否可以被改動

Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop

objIE.document.login.USERID.value=username
objIE.document.login.PASSWD.value=password

Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop


set WshShell = CreateObject("WScript.Shell")

WshShell.SendKeys "{enter}"
Wscript.Quit