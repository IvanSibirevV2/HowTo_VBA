'”рок VBScript є4:
'÷иклы While Е Wend и Do Е Loop
'file_5.vbs

dim i, arr_name(5), Arr_Filter, List

Arr_name(0) = "ѕонедельни "
Arr_name(1) = "¬торник"
Arr_name(2) = "—реда"
Arr_name(3) = "„етверг"
Arr_name(4) = "ѕ€тница"
Arr_name(5) = "—уббота"

Arr_Filter = Filter(Arr_name, "ик", false, 1)

i = 0

Do Until i = UBound(Arr_Filter) + 1
List = List & Arr_Filter(i) & vbCrlf
i = i+1
Loop

MsgBox List