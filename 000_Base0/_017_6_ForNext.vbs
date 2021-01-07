'Урок VBScript №6:
'Циклы For ... Next и For Each ... Next
'file_1.vbs
Dim test, i
'пример №2 без слова Step (шаг)
test = 0
For i=1 To 10
	test = test + i
next
MsgBox test
'пример №2 шаг положительный (2) идём от 1 до 10
test = 0
For i=1 To 10 Step 2
	test = test + i
next
MsgBox test
'пример №3 шаг отрицательный (-2) идём от 10 до 1
test = 0
For i=10 To 1 Step -2
	test = test + i
next
rem ForNext
MsgBox test