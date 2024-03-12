Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.Run"\\DNC10370-PC3\WinHotKey\VBS\auto\countdown.bat",0
Wscript.Sleep 50
WshShell.Run"\\DNC10370-PC3\WinHotKey\VBS\auto\auto攔截10分鐘強制登出SCMS.bat",0
Wscript.Sleep 50
WshShell.Run"\\DNC10370-PC3\WinHotKey\VBS\auto\監護輔助宣告.bat",0
Wscript.Sleep 50