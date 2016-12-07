
Sub GetFile(url, filename)
	dim xHttp: Set xHttp = CreateObject("MSXML2.ServerXMLHTTP")
	xHttp.Open "GET", url, False
	xHttp.setOption 2, 13056
	xHttp.Send
	Set aGet = CreateObject("ADODB.Stream")
	aGet.Mode = 3
	aGet.Type = 1
	aGet.Open() 
	aGet.Write(xHttp.responseBody)
	aGet.SaveToFile filename,2
End Sub


Set oArgs = WScript.Arguments
GetFile oArgs(0), oArgs(1)