@MODE CON COLS=35 LINES=5
@echo off
@color b0

CLS
ECHO.
ECHO.  ���b����Internet Explorer
ECHO.  Shut Down All Internet Explorer..
ECHO.
@taskkill /f /im iexplore.exe > nul

CLS
ECHO.
ECHO.  ���b����chrome
ECHO.  Shut Down All chrome..
ECHO.
@taskkill /f /im chrome.exe > nul

CLS
ECHO.
ECHO.  ���b�����Ҧ��O�ƥ�..
ECHO.  Shut Down All notepads..
ECHO.
@taskkill /f /im notepad.exe > nul

CLS
ECHO.
ECHO.  ���b����WINWORD
ECHO.  Shut Down All WINWORD..
ECHO.
@taskkill /f /im EXCEL.exe > nul






@taskkill /f /im wscript.exe > nul
@taskkill /f /im WINWORD.exe > nul

start file://dnc10370-pc3/WinHotKey/VBS/SCMS�`�γs��.vbs
@taskkill /f /im cmd.exe > nul
