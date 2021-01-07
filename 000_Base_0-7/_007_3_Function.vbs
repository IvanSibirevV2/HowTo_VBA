'Урок VBScript №3
'Функции (Function ... End Function) и процедуры (Sub ... End Sub)
'file_2.vbs

dim a, b

Sub sub_name(a, b)
	Dim rezult
	rezult = a+b
	MsgBox rezult
End Sub

Call sub_name(5,110)
Call sub_name(15,16)
Call sub_name(25,15)
sub_name 25,1





