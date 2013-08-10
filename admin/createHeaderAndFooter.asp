<%@ Language=VBScript%>
<%
option explicit

Dim strFileSource, strHeader, strFooter, objHeaderFile, objFooterFile, fso, strMiddle, objMiddleFile
Dim strReplaceWith, i
Function getFile(page)
	'***********************************************
	'* Inputs: String(page path - relative)
	'***********************************************
	'* Outputs: String(File Source Code)
	'***********************************************
	If CStr(page) <> "" Then
		Dim fsoViewSource, act, read_text
		set fsoViewSource = createobject("scripting.filesystemobject")
		set act = fsoViewSource.Opentextfile(Server.MapPath(page))
		read_text = act.readall
		act.Close
		Set fsoViewSource = Nothing
		getFile = read_text
	Else
		getFile = ""
	End If
End Function
strFileSource = getFile("/includes/default.asp")
For i=1 to 5
	strFileSource = Replace(strFileSource, VbCrLf & VbCrLf, VbCrLf)
Next

set fso = createobject("scripting.filesystemobject")

strHeader = Left(strFileSource, InStr(1, strFileSource, "<!---->")-4)
strFooter = Right(strFileSource, Len(strFileSource)-InStrRev(strFileSource, "<!---->")-7)

Set objHeaderFile = fso.CreateTextFile (Server.MapPath("/includes/header.asp"), True)
strHeader = Replace(strHeader, "href=""../", "href=""/")
strHeader = Replace(strHeader, "src=""../", "src=""/")
strHeader = Replace(strHeader, "background=""../", "background=""/")
strHeader = Replace(strHeader, "url('../", "url('/")
objHeaderFile.WriteLine(strHeader)
objHeaderFile.Close

Set objFooterFile = fso.CreateTextFile (Server.MapPath("/includes/footer.asp"), True)
strFooter = Replace(strFooter, "href=""../", "href=""/")
strFooter = Replace(strFooter, "src=""../", "src=""/")
strFooter = Replace(strFooter, "background=""../", "background=""/")
strFooter = Replace(strFooter, "url('../", "url('/")
objFooterFile.WriteLine(strFooter)
objFooterFile.Close

set fso = Nothing

Response.Redirect("/")
%>
