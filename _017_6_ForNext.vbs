'���� VBScript �6:
'����� For ... Next � For Each ... Next
'file_1.vbs
Dim test, i
'������ �2 ��� ����� Step (���)
test = 0
For i=1 To 10
	test = test + i
next
MsgBox test
'������ �2 ��� ������������� (2) ��� �� 1 �� 10
test = 0
For i=1 To 10 Step 2
	test = test + i
next
MsgBox test
'������ �3 ��� ������������� (-2) ��� �� 10 �� 1
test = 0
For i=10 To 1 Step -2
	test = test + i
next
rem ForNext
MsgBox test