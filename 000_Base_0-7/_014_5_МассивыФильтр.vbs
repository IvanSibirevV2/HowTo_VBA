'���� VBScript �4:
'����� While � Wend � Do � Loop
'file_5.vbs

Dim fun_name = Function(a, b)
begin
	Dim rezult
	rezult = a+b
	fun_name = rezult
	'����������� �������� �������, ������� ��� ��� �����
end

MsgBox fun_name(1,2)

rem Function fun_name(a, b)
rem      Dim rezult
rem      rezult = a+b
rem      fun_name = rezult
rem      '����������� �������� �������, ������� ��� ��� �����
rem End Function