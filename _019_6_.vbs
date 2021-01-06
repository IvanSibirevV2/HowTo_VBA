'Урок VBScript №6:
'Циклы For ... Next и For Each ... Next
'file_2.vbs

Dim arr_name(7),i,list

arr_name(0) = 45
arr_name(1) = 3.6
arr_name(2) = "строка"
arr_name(3) = #12.05.15#
arr_name(4) = #12/05/15#
arr_name(5) = 5*6
arr_name(6) = 5 - 1
arr_name(7) = -6

For Each i In arr_name     
     list = list & i & vbCrLf  
Next

MsgBox list