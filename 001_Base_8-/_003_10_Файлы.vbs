'********************************************************
'Урок VBScript №10:
'Работа с текстовыми документами - TextStream.
'file_1.vbs
'********************************************************

Dim fso, f1, ts, s
Const ForWriting = 2

Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateTextFile ("C:\D\test1.txt")
Set f1 = fso.GetFile("C:\D\test1.txt")
Set ts = f1.OpenAsTextStream(ForWriting, True)
ts.WriteLine("Тестирование 1, 2, 3.")
ts.Close
''''''''''''''''''''''''''''''
rem OpenTextFile(filename [,iomode [,create [,format]]])
rem filename - имя файла который необходимо прочитать
rem iomode - режим открытия файла (1 - только чтение, 2 - для записи (если уже существует, будет перезаписан), 8 - для добавления)
rem create - true - создать файл, если он не существует, false - не создавать
rem format - кодировка (-2 - кодировка ОС по умолчанию, -1 - Unicode, 0 - ASCII)
rem Непосредственно для чтения самих данных из текстового файла можно воспользоваться одним из следующих операторов:
rem Read - чтение определенного количества символов независимо от конца строки (количество символов указывается в скобках после оператора)
rem ReadLine - прочитать строку полностью до конца (до символов перевода строки), т.е. построчное чтение текстового файла
rem ReadAll - прочитать весь файл целиком за раз, включая символы переноса строки.
rem Эти же приемы чтения файла можно использовать и в следующем способе.
''''''''''''''''''''''''''''''
Set f = fso.OpenTextFile("C:\D\test1.txt", 1)
Do While Not f.AtEndOfStream
  str = f.ReadLine
  MsgBox str
Loop
f.Close


