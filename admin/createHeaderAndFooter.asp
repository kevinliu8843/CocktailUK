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
strReplaceWith = "<META name=""author"" content=""Lee Tracey"">" & VbCrLf & "<META name=""description"" content=""Cocktails are frivolous and fun, and are the perfect way to entertain with maximum style and minimum effort. We have easy to follow instructions for making devilishly delicious cocktails."">" & VbCrLf & "<META name=""keywords"" content=""cocktail recipe, cocktail drink recipe, cocktail bar drink recipe, recipe for cocktail, shooter, jello shooter, shooter game, shooter recipe, straight shooter, alcohol recipe, alcohol drink recipe, alcohol beverage recipe, alcohol punch recipe, recipe for making alcohol, drink recipe, mixed drink recipe, alcoholic drink recipe, recipe for mixed drink, alcohol drink recipe, recipe for alcoholic drink, alcoholic drink, alcoholic drink recipe, mixed alcoholic drink, recipe for alcoholic drink, alcoholic mixed drink, alcoholic drink recipe, alcoholic beverage recipe, recipe for alcoholic drink, alcoholic drink recipe mix, alcoholic cocktail recipe"">" & VbCrLf & "<MET  name=""revisit-after"" content=""30 days"">" & VbCrLf & "<META name=""robots"" content=""ALL"">" & VbCrLf & "<META name=""distribution"" content=""GLOBAL"">" & VbCrLf & "<META name=""rating"" content=""ADULTS"">" & VbCrLf & "<LINK REL=""SHORTCUT ICON"" href=""/favicon.ico"">" & VbCrLf

set fso = createobject("scripting.filesystemobject")

strHeader = Left(strFileSource, InStr(1, strFileSource, "<!---->")-4)
strFooter = Right(strFileSource, Len(strFileSource)-InStrRev(strFileSource, "<!---->")-7)

Set objHeaderFile = fso.CreateTextFile (Server.MapPath("/includes/header.asp"), True)
strHeader = Replace(strHeader, "<META name=""GENERATOR"" content=""Microsoft FrontPage 4.0"">"&VbCrLf, "", 1, -1, 1)
strHeader = Replace(strHeader, "<META name=""ProgId"" content=""FrontPage.Editor.Document"">"&VbCrLf, strReplaceWith, 1, -1, 1)
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
