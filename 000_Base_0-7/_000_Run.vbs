dim arr_name

arr_name = Array(Array(10,"строка",#12/05/2015#,54),Array(11,3,15,44),Array(18,37,16,34)) 'задаём значения для элементов массива
MsgBox arr_name(0)(2) & vbCrlf & arr_name(1)(2) & vbCrlf & arr_name(2)(2)