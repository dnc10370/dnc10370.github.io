Dim username,password,IE  ' 定義變量

'---------------------------------------------------------------網頁縮放大小100%
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.RegWrite"HKCU\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"
'---------------------------------------------------------------對應網站
StartURL ="https://xws.tstartel.com/xws/scmsLoginXws.do?token=NI2Tgi7GPZ7P3HPKEIAWvEfSXvsc8iTL0g6jtnysoDmsSAdlqqmVI6WKnGPwXlWxPSdhs4VznFEniNBqRh1TRa%2B38tjwSEsSKfZEsCeT99B7Yz9vB7HAlbQi2sjmeXkvBEnDnDYKCH4QP1mC1RB2ZS2E3YjytU0zwunE9zJfhvcHI8wNNwB1JKQOobT2leptqJKpf5HabbI695qbJlPyXoeBxtL4EA3VmBrWcuDPu9Tezlj4sqfdtcdYrVnU9EoqOjebn3Gac%2BxGE2vgw0Kke%2B5ium6Mmqp5atUbz2b%2Fy9s%3D&channelType=2"
'---------------------------------------------------------------指令開始執行
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE_" )
objIE.Visible = 1
objShell.AppActivate objIE
objIE.Navigate StartURL  '設定IE物件預設指向的頁面 
objIE.FullScreen=0      '全屏化IE物件
objIE.menubar=0         '不顯示IE物件選單欄
objIE.AddressBar=0       '不顯示IE物件位址列
objIE.ToolBar=0        '不顯示IE物件工具欄
objIE.StatusBar=0       '不顯示IE物件狀態列
objIE.Width=10         '設定IE物件寬度
objIE.Height=10         '設定IE物件高度
objIE.top=5         
objIE.left=5
objIE.resizable=0       '設定IE物件大小是否可以被改動
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop

WScript.Sleep 1500
Do while objIE.ReadyState<> 4 or objIE.busy    ' 用循環語句確保網頁加載完畢才執行下面操作
loop


'---------------------------------------------------------------對應網站
StartURL ="https://xws.tstartel.com/xws/initialBatchPay.do"
'---------------------------------------------------------------指令開始執行
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE_" )
objIE.Visible = 1
objShell.AppActivate objIE
objIE.Navigate StartURL  '設定IE物件預設指向的頁面 
objIE.FullScreen=0      '全屏化IE物件
objIE.menubar=0         '不顯示IE物件選單欄
objIE.AddressBar=0       '不顯示IE物件位址列
objIE.ToolBar=0        '不顯示IE物件工具欄
objIE.StatusBar=0       '不顯示IE物件狀態列
objIE.Width=900         '設定IE物件寬度
objIE.Height=500         '設定IE物件高度
objIE.top=0         
objIE.left=0
objIE.resizable=1       '設定IE物件大小是否可以被改動
