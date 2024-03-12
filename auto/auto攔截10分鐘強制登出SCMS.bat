@echo off
Title countdown
:start
start/wait \\DNC10370-PC3\WinHotKey\VBS\auto\auto攔截10分鐘強制登出SCMS.vbs
choice /t 0 /d y /n >nul
GOTO start
