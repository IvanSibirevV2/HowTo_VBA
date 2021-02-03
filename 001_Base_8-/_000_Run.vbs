Set WShell = CreateObject("WScript.Shell")
vbsInterpreter = "cscript.exe"
If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
    WShell.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
    WScript.Quit
End If
rem ####################################################
Randomize
dim GetRandomFileName_Check
dim GetRandomFileName_LastCheck
Function GetRandomFileName()
	Do
		GetRandomFileName_LastCheck = GetRandomFileName_Check
		GetRandomFileName_Check = _
			CreateObject("wscript.network").UserName &"_"& _
			DatePart("yyyy",Now)&"_"&_
			DatePart("m",Now)&"_"&_
			DatePart("d",Now)&"__"&_
			DatePart("h",Now)&"_"&_
			DatePart("n",Now)&"_"&_
			DatePart("s",Now)&"__"&_
			Int((100 * Rnd) + 1)&_
			".bat"    
	loop Until (GetRandomFileName_Check = GetRandomFileName_LastCheck)
    GetRandomFileName = GetRandomFileName_Check
End Function
WScript.StdOut.WriteLine "#################################"
i=0
While Not i = 20
    WScript.StdOut.WriteLine GetRandomFileName()
    i = i + 1
Wend
WScript.Sleep Int(9000)





