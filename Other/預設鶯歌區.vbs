on error resume next
StartURL ="https://scms.tcc.net.tw/ExtraApp/convergenceAddrQuery.do?module=EXTRA"
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
objIE.Width=700          '�]�wIE����e��
objIE.Height=400         '�]�wIE���󰪫�
objIE.top=252
objIE.left=280
objIE.resizable=1        '�]�wIE����j�p�O�_�i�H�Q���
Do while objIE.ReadyState<> 4 or objIE.busy    ' �δ`���y�y�T�O�����[�������~����U���ާ@
loop
Do while objIE.ReadyState<> 4 or objIE.busy    ' �δ`���y�y�T�O�����[�������~����U���ާ@
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


