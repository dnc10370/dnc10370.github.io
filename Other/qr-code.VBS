on error resume next
Dim username,password,IE,dom,form,Shell
Set fso = CreateObject("scripting.filesystemobject")
path="\\DNC10370-PC5\Public\IMEI\IMEI.txt"
Set A=fso.OpenTextFile(path,1)
For i=1 To 0
A.ReadLine
Next
line=A.ReadLine
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.RegWrite"HKCU\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",90000,"REG_DWORD"
StartURL ="\\DNC10370-PC3\WinHotKey\VBS\qr-code.html"
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
objIE.Width=80         '設定IE物件寬度
objIE.Height=265         '設定IE物件高度
objIE.top=465
objIE.left=20
objIE.resizable=0        '設定IE物件大小是否可以被改動

Do while objIE.ReadyState<> 4 or objIE.busy
loop
objIE.document.diyForm.content.value=line
objIE.document.diyForm.generate.click
Do while objIE.ReadyState<> 4 or objIE.busy
loop
Wscript.Quit