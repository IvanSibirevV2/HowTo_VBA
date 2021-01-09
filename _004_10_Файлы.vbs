'**********************************************
' fso createtextfile
' как создать текстовый файл
' запись данных в поток
' createtextfile.vbs
'**********************************************

Option Explicit
dim fso, c_file, t_file, WshShell, data, data_file  
' создаем ссылку на объект Wscript.Shell
set WshShell = CreateObject("Wscript.Shell")
' создаем ссылку на объект Scripting.FileSystemObject
set fso=CreateObject ("Scripting.FileSystemObject")

'**********************************************
' открываем текущий сценарий и записываем
' его содержимое в переменную data
'**********************************************

set c_file=fso.OpenTextFile(WScript.ScriptFullName, 1, false)
data = c_file.ReadAll
c_file.Close()

'------------------------------------------------------------------------------------ 
