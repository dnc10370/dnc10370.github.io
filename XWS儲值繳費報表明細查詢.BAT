@echo off
%1(start /min cmd.exe /c %0 :&exit)

echo msgbox "轉導至 XWS 中......",64,"轉導至...">alert.vbs && start alert.vbs && ping -n 2 127.1>nul && del alert.vbs
@ping 127.0.0.1 -n 2 -w 10000 > nul
start %~dp0\XWS儲值繳費報表明細查詢.vbs
@ping 127.0.0.1 -n 2 -w 10000 > nul

Taskkill /fi "windowTitle eq 轉導至..." /T
choice /t 0 /d y /n >nul

