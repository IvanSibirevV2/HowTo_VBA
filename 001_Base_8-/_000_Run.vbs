'�������� �������
Set WShell = CreateObject("WScript.Shell")
vbsInterpreter = "cscript.exe"
'����� �� �������
If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
    WShell.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
    WScript.Quit
End If
WScript.StdOut.WriteLine "##########################"
rem ���� ���� �������� ���������� �����.
rem ����� ������ ��������� �� ��������� � � ������ �� ���������
rem ���� �� ����������� �� N++
rem ����� �������� �� ����������
WShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject") _
	.GetParentFolderName(WScript.ScriptFullName)
If CreateObject("Scripting.FileSystemObject").FolderExists("KoteCombo") then
WScript.StdOut.WriteLine ("KoteCombo ������ �� ���")
else
WScript.StdOut.WriteLine ("KoteCombo ����� ��������")
set new_folder = CreateObject("Scripting.FileSystemObject").CreateFolder("KoteCombo")
End If
WScript.Sleep Int(4000)





