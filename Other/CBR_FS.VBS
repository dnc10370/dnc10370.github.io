on error resume next
Dim username,password,IE  ' 定義變量
Txt_Account = "DNC10370"  ' 輸入用戶名
Txt_Password = "Aa09350935"  ' 輸入密碼
StartURL ="http://cbr.tcc.net.tw/CBR_FS/Login.aspx"
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE_" )
objIE.Visible = 1
objShell.AppActivate objIE
objIE.Navigate StartURL  '設定IE物件預設指向的頁面 
objIE.FullScreen=0      '全屏化IE物件
objIE.menubar=0         '不顯示IE物件選單欄
objIE.AddressBar=0       '不顯示IE物件位址列
objIE.ToolBar=0         '不顯示IE物件工具欄
objIE.StatusBar=0       '不顯示IE物件狀態列
objIE.Width=1035         '設定IE物件寬度
objIE.Height=740         '設定IE物件高度
objIE.top=0
objIE.left=0
objIE.resizable=1       '設定IE物件大小是否可以被改動
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop
objIE.document.getelementByid("Txt_Account").value=Txt_Account   ' 修改（傳入）賬戶
objIE.document.getelementByid("Txt_Password").value=Txt_Password   ' 修改（傳入）密碼
objIE.document.getelementByid("Btn_loginIn").click          ' 模擬第一次登陸按鍵的點擊
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop



