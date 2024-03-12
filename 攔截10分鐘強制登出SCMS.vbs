Response = MsgBox("攔截10分鐘後強制登出SCMS 是否繼續?",vbYesNo)
if Response = vbYes Then
Response = MsgBox("已啟用", vbInformation , "攔截強制登出SCMS")
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.Run"\\DNC10370-PC3\WinHotKey\VBS\Other\auto.vbs"
Elseif Response =vbNo Then
Set objword = CreateObject("word.Application")
Set colTasks = objWord.Tasks
For i= 0 to 2
for each colTask in colTasks
if instr(colTask.name,"countdown")<>0 then
colTask.close
end if
next
next
Set WshShell =WScript.CreateObject("WScript.Shell")

WshShell.Run"taskkill /f /im wscript.exe",0
WshShell.Run"taskkill /f /im WINWORD.exe",0

End if