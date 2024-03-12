@echo off
@taskkill /f /im WINWORD.exe > nul
@ping 127.0.0.1 -n 1 -w 1000 > nul
@taskkill /f /im wscript.exe > nul
@Taskkill /F /FI "windowTitle eq countdown" /T
@ping 127.0.0.1 -n 1 -w 1000 > nul
@ping 127.0.0.1 -n 5 -w 1000 > nul

\\DNC10370-PC3\WinHotKey\VBS\Other\auto.VBS


