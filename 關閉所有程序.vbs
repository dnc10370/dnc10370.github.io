Response = MsgBox("�����Ҧ��{�� �O�_�~��?",vbYesNo)
If Response = vbYes Then
 MyString = "Yes"
Set WshShell =WScript.CreateObject("WScript.Shell")
WshShell.Run"\\DNC10370-PC3\WinHotKey\VBS\�����Ҧ��{��.bat"
Set WshShell = Nothing
Else
 MyString = "No"
End if