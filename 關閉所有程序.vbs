Response = MsgBox("關閉所有程序 是否繼續?",vbYesNo)
If Response = vbYes Then
 MyString = "Yes"
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.Run"\\DNC10370-PC3\WinHotKey\VBS\關閉所有程序.bat"
Set WshShell = Nothing
Else
 MyString = "No"
End if