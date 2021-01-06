'Урок VBScript №5:
'Массивы и фнкции для работы с ними
'file_7.vbs

Dim arr_name(3)

Function fun_name(a,b)
	fun_name = a+b
End Function

arr_name(0) = 100
arr_name(1) = "строка"
arr_name(2) = #12/05/2015#
Set arr_name(3) = getref("fun_name")

MsgBox arr_name(1)
MsgBox arr_name(3)(10,5)