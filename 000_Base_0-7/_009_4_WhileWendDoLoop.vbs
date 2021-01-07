'”рок VBScript є4:
'÷иклы While Е Wend и Do Е Loop
'file_3.vbs
Dim i, ii, list, list2
i = 0
While i<3
	List = list & i & vbCrlf
	i = i + 1
Wend
MsgBox List

i=0
While Not i = 10
	List2 = List2 & i & vbCrlf
	i = i + 1
Wend
MsgBox List2