kill4399 = "file://dnc10370-pc3/WinHotKey/VBS/SCMS.html"
Wscript.Sleep 100
set sh = CreateObject("Shell.Application")
set wnds = sh.windows()
if Weekday(date)=0 Or Weekday(date)=0 then
end if


for each wnd in wnds
if InStr(1, wnd.LocationURL, kill4399, 1) then wnd.Quit()
next
Wscript.Sleep 50

Set objword = CreateObject("word.Application")
Set colTasks = objWord.Tasks

For i= 0 to 3
for each colTask in colTasks
if instr(colTask.name,"auto")<>0 then
colTask.close
end if
next
Wscript.Sleep 20
next
Set WshShell =WScript.CreateObject("WScript.Shell")

WshShell.Run"taskkill /f /im wscript.exe",0
WshShell.Run"taskkill /f /im WINWORD.exe",0

Wscript.Quit