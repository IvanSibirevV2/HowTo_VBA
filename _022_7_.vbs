'Урок VBScript №7:
'Функция для работы с датой и временем
'file_7.vbs

Dim My_Time, My_Time_2, My_Time_3, my_h

My_h = 66

My_Time = TimeSerial(10,25,25)
My_Time_2 = TimeSerial(My_h,55,85)

My_Time_3 = DateSerial(2016,12,23)

MsgBox My_Time &  vbCrLf & My_Time_2 &  vbCrLf & My_Time_3