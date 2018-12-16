Dim xHttp: Set xHttp = CreateObject("Microsoft.XMLHTTP")
Dim bStrm: Set bStrm = CreateObject("ADODB.Stream")
Dim sStr:  Set sStr  = CreateObject("System.Text.String")

For part = 1 To 52
	xHttp.Open "GET", "http://85.25.117.95/mp3/1323/" + CStr(part) + ".mp3", False
	xHttp.Send
	With bStrm
		.type = 1 '//binary
		.open
		.write xHttp.responseBody
		.savetofile Right("0"+ CStr(part), 2) + ".mp3", 2 '//overwrite
		' .savetofile sStr.Format("{0:C2}.mp3"), 2 '//overwrite
		.close
	End With
Next
