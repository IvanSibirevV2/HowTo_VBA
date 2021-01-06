'Урок VBScript №7:
'Функция для работы с датой и временем
'file_6.vbs

Dim old_time, my_timer, i

old_time = Timer

i = 0
While i<10000000
	i = i + 1
Wend

my_timer = Timer - old_time 
MsgBox "Цикл был выполнен за " & my_timer & " секунд."