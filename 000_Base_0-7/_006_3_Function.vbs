'Урок VBScript №3
'Функции (Function ... End Function) и процедуры (Sub ... End Sub)
'file_1.vbs
 
 
dim a, b
 
Function fun_name(a, b)
	Dim rezult
	rezult = a+b
	fun_name = rezult 'присваиваем значение функцие, которое она нам вернёт
End Function
 
MsgBox fun_name(5,110)
MsgBox fun_name(15,16)
MsgBox fun_name(25,15)