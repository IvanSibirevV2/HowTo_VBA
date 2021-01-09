'********************************************************
'���� VBScript �10:
'������ � ���������� ����������� - TextStream.
'file_1.vbs
'********************************************************

Dim fso, f1, ts, s
Const ForWriting = 2

Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile ("C:\D\test1.txt")
Set f1 = fso.GetFile("C:\D\test1.txt")
Set ts = f1.OpenAsTextStream(ForWriting, True)
ts.WriteLine("������������ 1, 2, 3.")
ts.Close
''''''''''''''''''''''''''''''
rem OpenTextFile(filename [,iomode [,create [,format]]])
rem filename - ��� ����� ������� ���������� ���������
rem iomode - ����� �������� ����� (1 - ������ ������, 2 - ��� ������ (���� ��� ����������, ����� �����������), 8 - ��� ����������)
rem create - true - ������� ����, ���� �� �� ����������, false - �� ���������
rem format - ��������� (-2 - ��������� �� �� ���������, -1 - Unicode, 0 - ASCII)
rem ��������������� ��� ������ ����� ������ �� ���������� ����� ����� ��������������� ����� �� ��������� ����������:
rem Read - ������ ������������� ���������� �������� ���������� �� ����� ������ (���������� �������� ����������� � ������� ����� ���������)
rem ReadLine - ��������� ������ ��������� �� ����� (�� �������� �������� ������), �.�. ���������� ������ ���������� �����
rem ReadAll - ��������� ���� ���� ������� �� ���, ������� ������� �������� ������.
rem ��� �� ������ ������ ����� ����� ������������ � � ��������� �������.
''''''''''''''''''''''''''''''
Set f = fso.OpenTextFile("C:\D\test1.txt", 1)
Do While Not f.AtEndOfStream
  str = f.ReadLine
  MsgBox str
Loop
f.Close


