'Урок VBScript №2:
'Создание условий при помощи оператора select ... case
'file_3.vbs
Dim a
a = 3
Select case a
case 3
b = "b = 3"
case 10
b = "b = 10"
case else
b = "b = " & a
End Select
MsgBox b
