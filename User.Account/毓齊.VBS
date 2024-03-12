Dim username,password,IE  ' 定義變量
username = "C121615835"  ' 輸入用戶名
password = "Aa123456"  ' 輸入密碼
'---------------------------------------------------------------網頁縮放大小100%
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.RegWrite"HKCU\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"
'---------------------------------------------------------------對應網站
StartURL ="https://scms.tcc.net.tw/"
'---------------------------------------------------------------指令開始執行
on error resume next
Set WshShell =WScript.CreateObject("WScript.Shell")

WshShell.Run"Taskkill /fi ""windowTitle eq SCMS SSO - Internet Explorer"" /T",0

Wscript.Sleep 200

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
objIE.Width=1035         '設定IE物件寬度
objIE.Height=740         '設定IE物件高度
objIE.top=0         
objIE.left=0
objIE.resizable=1       '設定IE物件大小是否可以被改動

For i= 0 to 100
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop
objIE.document.getelementByid("username").value=username   ' 修改（傳入）賬戶
objIE.document.getelementByid("password").value=password   ' 修改（傳入）密碼
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop
Wscript.Sleep 100
next