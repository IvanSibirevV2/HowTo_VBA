'Создание консоли
Set WShell = CreateObject("WScript.Shell")
vbsInterpreter = "cscript.exe"
'Фокус на консоль
If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
    WShell.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
    WScript.Quit
End If
WScript.StdOut.WriteLine "##########################"
rem Есть один незримый вселенский косяк.
rem Точка старта программы НЕ совпадает с её папкой по умолчанию
rem Если Вы запускатете из N++
rem Косяк досадный но исправимый
WShell.CurrentDirectory = CreateObject("Scripting.FileSystemObject") _
	.GetParentFolderName(WScript.ScriptFullName)
If CreateObject("Scripting.FileSystemObject").FolderExists("KoteCombo") then
WScript.StdOut.WriteLine ("KoteCombo Создан до нас")
else
WScript.StdOut.WriteLine ("KoteCombo Будет сотворен")
set new_folder = CreateObject("Scripting.FileSystemObject").CreateFolder("KoteCombo")
End If
WScript.Sleep Int(4000)





