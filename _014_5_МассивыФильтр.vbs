'���� VBScript �4:
'����� While � Wend � Do � Loop
'file_5.vbs

dim i, arr_name(5), Arr_Filter, List

Arr_name(0) = "�����������"
Arr_name(1) = "�������"
Arr_name(2) = "�����"
Arr_name(3) = "�������"
Arr_name(4) = "�������"
Arr_name(5) = "�������"

Arr_Filter = Filter(Arr_name, "��", false, 1)

i = 0

Do Until i = UBound(Arr_Filter) + 1
List = List & Arr_Filter(i) & vbCrlf
i = i+1
Loop

MsgBox List