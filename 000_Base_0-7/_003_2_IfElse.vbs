'���� VBScript �2:
'�������� ������� ��� ������ ��������� if, else � ��� �� select, case
'file_1.vbs
 
Dim a, b, c
 
a = 10
b = 20
c = 40.5
 
'������ ������ �����������
if a = 10 then MsgBox "a = 10"
 
'������ ������ �����������
if b = 10 then
	MsgBox "b = 10"
else
	MsgBox "b �� ����� 10"
end if
 
'������ ������� �����������
if c = 10 then
	MsgBox "c = 10"
elseif c = 20 then
	MsgBox "b = 20"
elseif c = 40.5 then
	MsgBox "b = 40,5"
	MsgBox "���!"
else
	MsgBox "���������� �� �������"
End if
