rem arr=Split("w3ii is my favourite website")
rem for each index in arr
rem      List = List & index & vbCrLf
rem next
rem MsgBox List

arr=Split("�����������,�������,�����,�������,�������,�������",",")
for each index in arr
     List = List & index & vbCrLf
next
MsgBox List

rem arr = Array("�����������","�������","�����","�������","�������","�������")
rem Arr_Join = Join(arr, ",")
rem MsgBox Arr_Join