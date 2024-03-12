@echo off
%1(start /min cmd.exe /c %0 :&exit)
taskkill /f /im WINWORD.exe > nul
ping 127.0.0.1 -n 1 -w 1000 > nul
taskkill /f /im wscript.exe > nul

ping 127.0.0.1 -n 1 -w 1000 > nul

for /L %%i in (1, 1, 5) do (
Taskkill /F /fi "windowTitle eq countdown" /T
choice /t 0 /d y /n >nul
)


exit