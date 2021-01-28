rem arr=Split("w3ii is my favourite website")
rem for each index in arr
rem      List = List & index & vbCrLf
rem next
rem MsgBox List

arr=Split("ПонедельниК,Вторник,Среда,Четверг,Пятница,Суббота",",")
for each index in arr
     List = List & index & vbCrLf
next
MsgBox List

rem arr = Array("ПонедельниК","Вторник","Среда","Четверг","Пятница","Суббота")
rem Arr_Join = Join(arr, ",")
rem MsgBox Arr_Join