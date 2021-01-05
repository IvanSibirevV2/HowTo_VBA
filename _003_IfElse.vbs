'Урок VBScript №2:
'Создание условий при помощи оператора if, else а так же select, case
'file_1.vbs
 
Dim a, b, c
 
a = 10
b = 20
c = 40.5
 
'пример первой конструкции
if a = 10 then MsgBox "a = 10"
 
'пример второй конструкции
if b = 10 then
	MsgBox "b = 10"
else
	MsgBox "b не равно 10"
end if
 
'пример третьей конструкции
if c = 10 then
	MsgBox "c = 10"
elseif c = 20 then
	MsgBox "b = 20"
elseif c = 40.5 then
	MsgBox "b = 40,5"
	MsgBox "ура!"
else
	MsgBox "совпадений не найдено"
End if
