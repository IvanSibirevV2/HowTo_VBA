rem День независимости
rem День независимости, возрождение
MsgBox _
	"Ищем символ ...- смотрите,не найдем" &  vbCrlf & _
	"123456" &  vbCrlf & _
	"abcdef" &  vbCrlf & _
	"abc^" &  vbCrlf & _
	"InStrRev(""abcdef"",""D"",3)=" & InStrRev("abcdef","D",6,0) & vbCrlf & _
"!!!"