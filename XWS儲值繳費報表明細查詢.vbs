on error resume next
Dim username,password,IE  ' �w�q�ܶq
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.RegWrite"HKCU\Software\Microsoft\Internet Explorer\Zoom\ZoomFactor",100000,"REG_DWORD"



StartURL ="https://xwsrpt.tstartel.com/xws-report/initialDealerIncomeDeatil.do?user_id=DNC10370_H124196305&function_id=110001&log_time=2023/12/09 14:04:30&channel_id=002&collecting_fees=X&collecting_fees_name=X&collectin"
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "objIE_")
objIE.Visible = 1
objShell.AppActivate objIE
objIE.Navigate StartURL '�]�wIE����w�]���V������ 
objIE.FullScreen=0       '���̤�IE����
objIE.menubar=0          '�����IE��������
objIE.AddressBar=0       '�����IE�����}�C
objIE.ToolBar=0          '�����IE����u����
objIE.StatusBar=0        '�����IE���󪬺A�C
objIE.Width=700         '�]�wIE����e��
objIE.Height=500         '�]�wIE���󰪫�
objIE.Top=50
objIE.left=20
objIE.resizable=1        '�]�wIE����j�p�O�_�i�H�Q���


StartURL ="https://xwsrpt.tstartel.com/xws-report/initialDealerRechargeDetail.do?user_id=DNC10370_H124196305&function_id=110041&log_time=2023/12/09 14:04:30&channel_id=002&collecting_fees=X&collecting_fees_name=X&collecting_channel_id=X&collecting_channel_name=X"
Set objShell = WScript.CreateObject("WScript.Shell")
Set objIE = WScript.CreateObject("InternetExplorer.Application", "objIE_")
objIE.Visible = 1
objShell.AppActivate objIE
objIE.Navigate StartURL '�]�wIE����w�]���V������ 
objIE.FullScreen=0       '���̤�IE����
objIE.menubar=0          '�����IE��������
objIE.AddressBar=0       '�����IE�����}�C
objIE.ToolBar=0          '�����IE����u����
objIE.StatusBar=0        '�����IE���󪬺A�C
objIE.Width=700         '�]�wIE����e��
objIE.Height=500         '�]�wIE���󰪫�
objIE.Top=150
objIE.left=50
objIE.resizable=1        '�]�wIE����j�p�O�_�i�H�Q���