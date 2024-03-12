Response = MsgBox("關閉所有IE瀏覽器視窗 是否繼續?",vbYesNo)
If Response = vbYes Then
 MyString = "Yes"
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.Run"taskkill /f /im iexplore.exe",0
WshShell.Run"\\DNC10370-PC3\WinHotKey\VBS\Other\login-ap.vbs"
Set objword = CreateObject("word.Application")
For i= 0 to 10
Set colTasks = objWord.Tasks
for each colTask in colTasks
if instr(colTask.name,"countdown")<>0 then
colTask.close
end if
next
Wscript.Sleep 20
next
Else
 MyString = "No"
End if