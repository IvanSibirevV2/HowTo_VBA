'���� VBScript �7:
'������� ��� ������ � ����� � ��������
'file_7.vbs

Dim My_Time, My_Time_2, My_Time_3, my_h

My_h = 66
rem DateTimeSerial 
My_Time = TimeSerial(10,25,25)
My_Time_2 = TimeSerial(My_h,55,85)

My_Time_3 = DateSerial(2016,12,23)

MsgBox My_Time &  vbCrLf & My_Time_2 &  vbCrLf & My_Time_3

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
MsgBox WeekDay(#14/02/1995#, vbSunday)'???
MsgBox  WeekdayName(3, False, vbMonday) &  vbCrLf & WeekdayName(3, True, vbMonday)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'vbUseSystemDayOfWeek � ������������ ��������� �� ��������� �������
'vbSunday � �����������
'vbMonday � �����������
'vbTuesday -�������
'vbWednesday � �����
'vbThursday � �������
'vbFriday � �������
'vbSaturday � �������
'''''''''''''''''''''''''''''''''''''''''''''''''''''''