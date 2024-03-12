on error resume next
Set objword = CreateObject("word.Application")
Set colTasks = objWord.Tasks
for each colTask in colTasks
For i= 0 to 5
if instr(colTask.name,"countdown")<>0 then
colTask.close
end if
next
Wscript.Sleep 100
next

Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.Run"taskkill /f /im WINWORD.exe",0
WshShell.Run"taskkill /f /im wscript.exe",0

