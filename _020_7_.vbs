'Урок VBScript №7:
'Функция для работы с датой и временем
'file_1.vbs

Dim my_date, my_date1, my_date2, my_date3

my_date = Date
my_date1 = Now
my_date2 = #12/05/2010#
my_date3 = #15.05.10#

MsgBox "Date = " & my_date & vbCrLf &"Now = " & my_date1 & vbCrLf & my_date2 & vbCrLf & my_date3