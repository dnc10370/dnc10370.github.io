@echo off
Title countdown
:start
Taskkill /fi "windowTitle eq Windows Script Host" /T
choice /t 0 /d y /n >nul

goto start