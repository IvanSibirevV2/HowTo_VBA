rem Abs � ��������� �������� ���������� �����. ����� ������ � ������.
rem Atn � ������ ���������� �����.
rem Cos � ���������� ������� ����, ��������� � ��������. ������ ���������� �� -1 �� 1.
rem Sin � ������ ����� ����, ������� �� ������� � ��������. ������ ���������� �� -1 �� 1.
rem Tan � ���������� ������� ����. �������� ���� ��������� � ��������.
rem Exp � ���������� ��� ����������, ���������� � ��������� �����.
rem Log � �������� ����������� �������� �� ���������� �����. ����� ������ ���� ������ 0.
rem Sqr � ������ ���������� ������ ���������� �����.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MsgBox _
	"������ ����� 10: " & Abs(10) & vbCrLf & _
	"���������� ����� 10: " & Atn(10) & vbCrLf & _
	"������� ���� 10 ������: " & Cos(10) & vbCrLf & _
 	"����� ���� 10 ������: " & Sin(10) & vbCrLf  & _
	"������� ���� 10 ������: " & Tan(10) & vbCrLf  & _
	"���������� ���������� � 2: " & exp(2) & vbCrLf  & _
	"�������� �� 15: " & log(15) & vbCrLf  & _
	"���������� �������� �� 15: " & log(15)/log(10) & vbCrLf  & _
	"���������� ������ 15: " & Sqr(15) & vbCrLf  & _
	"���������� ������ 0: " & Sqr(0) & vbCrLf  & _
	"!!!"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Randomize
MsgBox _
	"��������� ����� � ��������� �� 0 �� 1" & vbCrLf  & _
	Rnd & vbCrLf  & Rnd & vbCrLf  & Rnd & vbCrLf  & _
	Rnd & vbCrLf  & Rnd & vbCrLf  & Rnd & vbCrLf  & _
	Rnd & vbCrLf  & Rnd & vbCrLf  & Rnd & vbCrLf  & _
	Rnd & vbCrLf  & Rnd & vbCrLf  & Rnd & vbCrLf  & _
	Rnd & vbCrLf  & Rnd & vbCrLf  & Rnd & vbCrLf  & _
	"!!!"
	rem Int((10 * Rnd) + 1) 
MsgBox _
	"��������� ����� ����� � ��������� �� 0 �� 10" & vbCrLf  & _
	Int((10 * Rnd) + 1) & vbCrLf  &_
	Int((10 * Rnd) + 1) & vbCrLf  &_
	Int((10 * Rnd) + 1) & vbCrLf  &_
	Int((10 * Rnd) + 1) & vbCrLf  &_
	Int((10 * Rnd) + 1) & vbCrLf  &_
	Int((10 * Rnd) + 1) & vbCrLf  &_
	"!!!"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MsgBox _
	"Int(10.3) = "& Int(10.3) & vbCrLf  &_
	"Fix(10.3) = "& Fix(10.3) & vbCrLf  &_
	"Int(-10.3) = "& Int(-10.3) & vbCrLf  &_
	"Fix(-10.3) = "& Fix(-10.3) & vbCrLf  &_
"!!!"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
rem FormatNumber
rem Expression � ������������ ������������ ��������. ���� �����.
rem Numdigits � ���������� ������ ����� ���������� ����� (�������). ���� -1, �� �� ���������� �������.
rem Leadingdigit � ����������, ����� �� ������������ ���� ����� �� �������, � ������� ���������� ������ ������� (00,1)
rem Useparens � ����� �� ��������� ������������� �������� � ������� ������.
rem Groupdigits � ����� �� ����������� ���� � �����.
MsgBox _
	"FormatNumber(-00.13,3,-1,-1,vbTrue) = "&FormatNumber(-00.13,3,-1,-1,vbTrue) & vbCrLf  &_
	"FormatNumber(-00.13,3,-1,0,0) = "&FormatNumber(-00.13,3,-1,0,0)& vbCrLf  &_
	"FormatNumber(-00.13) = "&FormatNumber(-00.13)& vbCrLf  &_
"!!!"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
rem FormatCurrency � ����������� ����� ��� �������� ������ � ���������� ��������� �����.
rem FormatPercent � ���������� ����������������� ����� ��� ��������. �������� �� 100 � ��������� ���� %.
MsgBox _
	"FormatCurrency(1000,3,-1,-1) = " & FormatCurrency(1000,3,-1,-1)& vbCrLf  &_
	"FormatPercent(100,0,-1,-1) = " & FormatPercent(100,0,-1,-1)& vbCrLf  &_
"!!!"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MsgBox _
	"Hex(1000) = " & "16x" & Hex(1000)& vbCrLf  &_
	"Oct(1000) = " & "8x"& Oct(1000)& vbCrLf  &_
"!!!"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MsgBox _
	"�������� �� �����" & vbCrLf  &_
	"IsNumeric(""67325"") = " & IsNumeric("67325")& vbCrLf  &_
	"IsNumeric(""673ae25"") = " & IsNumeric("673ae25")& vbCrLf  &_
	"IsNumeric(""fd,glkdfhgjk"") = " & IsNumeric("fd,glkdfhgjk")& vbCrLf  &_
"!!!"