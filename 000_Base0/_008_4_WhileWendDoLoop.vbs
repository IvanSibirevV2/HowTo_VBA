'Урок VBScript №4:
'Циклы While … Wend и Do … Loop
rem WhileWendDoLoop
'file_1.vbs
dim i
i = 0
Do While (i<3)
	MsgBox "Пример №1: " & i
	i = i + 1
loop
i = 0
Do
	MsgBox "Пример №2: " & i
	i = i + 1
loop  While (i<3)
i = 0
Do Until (i>3)
	MsgBox "Пример №3: " & i
	i = i + 1
loop
i = 0
Do
	MsgBox "Пример №4: " & i
	i = i + 1
loop  Until (i>3)
'Урок VBScript №4:
'Циклы While … Wend и Do … Loop
'file_2.vbs
rem dim i
i = 0
Do While (i<3)
	MsgBox "Пример №1: " & i
	i = i + 1
	If i = 2 Then Exit Do
loop
