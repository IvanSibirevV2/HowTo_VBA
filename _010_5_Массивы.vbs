'���� VBScript �5:
'������� � ������ ��� ������ � ����
'file_1.vbs

dim arr_name(2)            '������ �� ��� ���������
dim matrix_name(1,3)       '������� 1 (������) �� 4 (��������)
dim dyn_name()             '������������ ������

arr_name(0) = 100
arr_name(1) = "������"
arr_name(2) = #12/05/2015#

matrix_name(0,0) = 1
matrix_name(0,1) = 10
matrix_name(0,2) = 100
matrix_name(0,3) = 1000

matrix_name(1,0) = 2
matrix_name(1,1) = 20
matrix_name(1,2) = 200
matrix_name(1,3) = 2000

MsgBox arr_name(1) & VbCrLf & matrix_name(0,2) & VbCrLf & matrix_name(1,2)