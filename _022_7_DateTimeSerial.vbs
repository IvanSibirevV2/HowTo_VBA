'”рок VBScript є7:
'‘ункци€ дл€ работы с датой и временем
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
MsgBox DateValue(now) &  vbCrLf & DateValue("14, февраль, 1995") &  vbCrLf & DateValue("02,14,1995") &  vbCrLf & DateValue("02/14/1995")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
MsgBox MonthName(9,true) &  vbCrLf & MonthName(9,false)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
MsgBox " Now = " & Now & vbCrLf & "FormatDateTime(Now, 0) = " & FormatDateTime(Now, 0) & vbCrLf & "FormatDateTime(Now, 1) = " & FormatDateTime(Now, 1) & vbCrLf & "FormatDateTime(Now, 2) = " & FormatDateTime(Now, 2) & vbCrLf & "FormatDateTime(Now, 3) = " & FormatDateTime(Now, 3) & vbCrLf & "FormatDateTime(Now, 4) = " & FormatDateTime(Now, 4)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsDate("10/14/2012") then
	MsgBox "ƒата корректна!"
else
	MsgBox "ƒата не корректна!"
end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'vbUseSystemDayOfWeek Ч использовать настройки по умолчанию системы
'vbSunday Ч воскресенье
'vbMonday Ч понедельник
'vbTuesday -вторник
'vbWednesday Ч среда
'vbThursday Ч четверг
'vbFriday Ч п€тница
'vbSaturday Ч суббота
'''''''''''''''''''''''''''''''''''''''''''''''''''''''