on error resume next
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.RegWrite"HKCU\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"
StartURL ="\\DNC10370-PC3\WinHotKey\VBS\SCMS�`�γs��.html"
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "IE_")
objIE.Visible = 1
objShell.AppActivate objIE

objIE.Navigate StartURL  '�]�wIE����w�]���V������ 
objIE.FullScreen=0       '���̤�IE����
objIE.menubar=0          '�����IE��������
objIE.AddressBar=0       '�����IE�����}�C
objIE.ToolBar=0          '�����IE����u����
objIE.StatusBar=0        '�����IE���󪬺A�C
objIE.Width=445         '�]�wIE����e��
objIE.Height=515         '�]�wIE���󰪫�
objIE.top=120
objIE.left=285

objIE.resizable=0        '�]�wIE����j�p�O�_�i�H�Q���
Do while objIE.ReadyState<> 4 or objIE.busy    ' �δ`���y�y�T�O�����[�������~����U���ާ@
loop
objIE.document.getElementByid("login-ap").click  

Do while objIE.ReadyState<> 4 or objIE.busy    ' �δ`���y�y�T�O�����[�������~����U���ާ@
loop

Wscript.Quit