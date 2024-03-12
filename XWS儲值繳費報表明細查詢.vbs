on error resume next
Dim username,password,IE  ' 定義變量
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.RegWrite"HKCU\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"



StartURL ="https://xwsrpt.tstartel.com/xws-report/initialDealerIncomeDeatil.do?user_id=DNC10370_H124196305&function_id=110001&log_time=2023/12/09 14:04:30&channel_id=002&collecting_fees=X&collecting_fees_name=X&collectin"
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "objIE_")
objIE.Visible = 1
objShell.AppActivate objIE
objIE.Navigate StartURL '設定IE物件預設指向的頁面 
objIE.FullScreen=0       '全屏化IE物件
objIE.menubar=0          '不顯示IE物件選單欄
objIE.AddressBar=0       '不顯示IE物件位址列
objIE.ToolBar=0          '不顯示IE物件工具欄
objIE.StatusBar=0        '不顯示IE物件狀態列
objIE.Width=700         '設定IE物件寬度
objIE.Height=500         '設定IE物件高度
objIE.Top=50
objIE.left=20
objIE.resizable=1        '設定IE物件大小是否可以被改動


StartURL ="https://xwsrpt.tstartel.com/xws-report/initialDealerRechargeDetail.do?user_id=DNC10370_H124196305&function_id=110041&log_time=2023/12/09 14:04:30&channel_id=002&collecting_fees=X&collecting_fees_name=X&collecting_channel_id=X&collecting_channel_name=X"
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "objIE_")
objIE.Visible = 1
objShell.AppActivate objIE
objIE.Navigate StartURL '設定IE物件預設指向的頁面 
objIE.FullScreen=0       '全屏化IE物件
objIE.menubar=0          '不顯示IE物件選單欄
objIE.AddressBar=0       '不顯示IE物件位址列
objIE.ToolBar=0          '不顯示IE物件工具欄
objIE.StatusBar=0        '不顯示IE物件狀態列
objIE.Width=700         '設定IE物件寬度
objIE.Height=500         '設定IE物件高度
objIE.Top=150
objIE.left=50
objIE.resizable=1        '設定IE物件大小是否可以被改動