'Урок VBScript №1
Rem Знакомство с языком VBScript
'file_1.vbs
 Rem Кодировка в ANSI
 
 MsgBox "Привет мир"
 
 Dim a, b, c, d, f
 
a = 10
b = 20
c = 40
d = "пробная строка"
f = 25
 
MsgBox a & b & c & d & f
MsgBox a & vbTab & b & vbTab & c & vbTab & d & vbTab & f
MsgBox a & vbCrLf & b & vbCrLf & c & vbCrLf & d & vbCrLf & f