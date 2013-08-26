<!--#include virtual="/includes/rating.asp" -->
<%
Sub Do301Redirect(strNewURL)
    Response.Status = "301 Moved Permanently"
    Response.AddHeader "Location", strNewURL
    Response.End
End Sub

Function strIntoDB( strString )
	strString = Replace ( strString, Chr(39), Chr(39)&Chr(39) )
	strString = Replace ( strString, VbCrLf, "<br/>" )
	strString = Replace ( strString, Chr(13), "<br/>" ) 'need this aswell???
	strString = Replace ( strString, """", Chr(39)&Chr(39) )
	strIntoDB = strString
End Function

Function strOutDB( ByVal strString )
	strString = strString & ""
	If NOT IsNull(strString) Then
		strString = Trim(strString)
	End If
	strOutDB = strString
End Function

Function Min(argOne, argTwo)
	If IsNumeric(argOne) AND IsNUmeric(argTwo) Then
		If argOne > argTwo Then
			Min = argTwo
		Else
			Min = argOne
		End If
	Else
		Min = null
	End If
End Function

Function Max(argOne, argTwo)
	If IsNumeric(argOne) AND IsNUmeric(argTwo) Then
		If argOne < argTwo Then
			Max = argTwo
		Else
			Max = argOne
		End If
	Else
		Max = null
	End If
End Function

Function replaceStuff ( strString )
	strString = Replace ( strString, Chr(39), Chr(39)&Chr(39) )
	strString = Replace ( strString, Chr(13), "<BR>" )
	replaceStuff = strString
End Function

Function replaceStuffBack( strString )
	strString = Replace(strString, """", "")
	strString = Replace(strString, "''", "'")
	replaceStuffBack = strString
End Function

Function capitalise( strString )
	If Len(strString) > 0 Then
		capitalise = ( UCase(Left( strString, 1 )) & Right ( strString, Len (strString) - 1 ) )
	End If
End Function

Sub displayPageLocation(strTitle, strTitleOut, strTopTitle, strUrl, strLinkStyle)
	Dim strUrl2, cArrLinks(10), cArrText(10), i, strTitleMod, ipos, ilen, blnShop, blnProduct, strTopTitleStore

	strTopTitleStore = strTopTitle
	blnShop = (InStr(Request.ServerVariables("SCRIPT_NAME"), "/shop/") > 0)
	blnProduct = (InStr(Request.ServerVariables("SCRIPT_NAME"), "/shop/products/product.asp") > 0)

	cArrLinks(0) = "/"
	cArrText(0) = "Cocktail : UK"
	
	If checkUrl(strUrl, "basedon.asp") Then
		cArrLinks(1) = "/cocktails/basedon.asp"
		cArrText(1) = "Cocktails Based On..."
	End If
	If checkUrl(strUrl, "/search") OR checkUrl(strUrl, "cocktails/containing.asp") Then
		cArrLinks(1) = "/search"
		cArrText(1) = "Search"
	End If
	If checkUrl(strUrl, "/account/submitCocktail.asp") Then
		cArrLinks(2) = "/account/submitcocktail.asp"
		cArrText(2) = "Submit drink"
	End If
	If checkUrl(strUrl, "/account") Then
		cArrLinks(1) = "/account/login.asp"
		cArrText(1) = "Members Area"
	End If
	If checkUrl(strUrl, "/db/stats") Then
		cArrLinks(1) = "/db/stats"
		cArrText(1) = "Top Ten..."
	End If
	If checkUrl(strUrl, "/features") Then
		cArrLinks(1) = "/features"
		cArrText(1) = "Features"
	End If
	If checkUrl(strUrl, "/features/media") Then
		cArrLinks(2) = "/features/media"
		cArrText(2) = "Media Releases"
	End If
	If checkUrl(strUrl, "/bartending/") Then
		cArrLinks(2) = "/features/bartending"
		cArrText(2) = "Bartending"
	End If
	If checkUrl(strUrl, "/webmaster") Then
		cArrLinks(1) = "/webmaster"
		cArrText(1) = "Webmaster Section"
	End If
	If checkUrl(strUrl, "/flair") Then
		cArrLinks(1) = "/flair"
		cArrText(1) = "Bartending Courses"
	End If
	If checkUrl(strUrl, "/services") Then
		cArrLinks(1) = "/services"
		cArrText(1) = "Services"
	End If
	If checkUrl(strUrl, "/admin") Then
		cArrLinks(1) = "/"
		cArrText(1) = "Administration Section"
	End If
	If checkUrl(strUrl, "/shop") Then
		cArrLinks(1) = "/shop"
		cArrText(1) = "Bar Equipment Shop"
	End If
	If checkUrl(strUrl, "/articles") Then
		cArrLinks(1) = "/articles"
		cArrText(1) = "Articles"
	End If
	If checkUrl(strUrl, "/competition") Then
		cArrLinks(1) = "/competition"
		cArrText(1) = "Competitions"
	End If
	If checkUrl(strUrl, "/db/link") Then
		cArrLinks(1) = "/db/link"
		cArrText(1) = "Links"
	End If
	
	strUrl = ""
	For i=1 to UBound(cArrLinks)
		If cArrLinks(i-1) <> "" Then
			strUrl2 = strUrl2 & "<A HREF=""" & cArrLinks(i-1) & """ style="""&strLinkStyle&""" class=""nocolour"">" & cArrText(i-1) & "</A> > "
			strTopTitle = strTopTitle & cArrText(i-1) & " > "
		Else
			Exit For
		End If
	Next
	
	'Now handled by ellipsis style tag (not on all browsers tho :( )
	If Request.ServerVariables("URL") <> "/" AND LCase(Request.ServerVariables("URL")) <> "/default.asp" Then
		strTopTitle = strTopTitle & strTitle
		strTitleMod = strTitle
		ilen = 40
		if Len(strTitle) > ilen then
			strTitleMod = Left(strTitleMod, ilen) & "..."
		end if
		strTitleOut = strUrl2 & strTitleMod
	End If
	
	If blnHardwireTitle then
		strTopTitle = strTitle
	End if
	
	If blnShop then
		strTopTitle = strTopTitleStore
	End if
	
	If strTitle = "" Then
		strTopTitle = "Cocktail : UK - cocktail recipes,bar equipment,shooters,home bars,drink recipes,bar equipment,cocktails"
	End If
End Sub

Function checkUrl(strUrl, strToCheck)
	checkUrl = (InStr(LCase(strUrl), LCase(strToCheck)) > 0)
End Function

Function hasImageThumb( name )
	Dim FSO
	Set FSO = Server.CreateObject ("scripting.filesystemobject")
	name = Trim( Replace( replaceStuffBack( name ), ",", "" ))

	IF fso.FileExists(Server.mappath("/images/cocktailThumbs/"& name &".jpg" ) ) THEN
		hasImageThumb= TRUE
	ELSE
		hasImageThumb= FALSE
	END IF
	Set FSO = Nothing
End Function 

Private Function SendEmail(strFrom, strTo, strCC, strBcc, strSubjectIn, strBody, blnHTML, strAttachment)
	Dim blnCanClearError
	
	On Error Resume Next
	SendEmail = True

	Dim Mail, strSubject
	strSubject = Replace(strSubjectIn, "&amp;" , "&")
	Const cdoSendUsingMethod        = "http://schemas.microsoft.com/cdo/configuration/sendusing"
	Const cdoSMTPServer             = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
	Const cdoSMTPServerPort         = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
	Const cdoSMTPConnectionTimeout  = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
	Const cdoSMTPAuthenticate       = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
	Const cdoSendUserName           = "http://schemas.microsoft.com/cdo/configuration/sendusername"
	Const cdoSendPassword           = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
	Const cdoMailboxURL 		= "http://schemas.microsoft.com/cdo/configuration/mailboxurl"
	Const cdoPickupDirectory	= "http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory"
	Const cdoSendUsingPickup	= 1
	Const cdoSendUsingPort		= 2
	Const cdoSendUsingExchange	= 3
	Const cdoNone			= 0
	Const cdoBasic                  = 1
	Const cdoNTLM			= 2
	Const pickupfolder		= "C:\ExchangePickup\"

	Dim objConfig
	Dim objMessage
	Dim Fields
	
	' Get a handle on the config object and it's fields
	Set objConfig = Server.CreateObject("CDO.Configuration")
	Set Fields = objConfig.Fields
	
	' Set config fields we care about
	With Fields
		.Item(cdoSendUsingMethod)       = cdoSendUsingPickup
		.Item(cdoPickupDirectory)       = pickupfolder
	'	.Item(cdoSMTPServer)            = "exchange.2tlimited.com"
	'	.Item(cdoSMTPServerPort)        = 25
	'	.Item(cdoSMTPConnectionTimeout) = 10
   	'	.Item(cdoSMTPAuthenticate) 	= cdoBasic
	'	.Item(cdoSendUsername) 		= "lee"
	'	.Item(cdoSendPassword) 		= "Smetsy#149"
		.Update
	End With

	Set objMessage = Server.CreateObject("CDO.Message")
	Set objMessage.Configuration = objConfig
	With objMessage
		.To       = strTo
		.From = strFrom
	'	.From     = "theteam@cocktail.uk.com"
	'	.Headers.Add "Reply-To", strFrom
		If strCC <> "" Then
			.CC = strCC
		End If
		If strBCC <> "" Then
			.BCC = strBCC
		End If


		.Subject  = strSubject
		If blnHTML Then
			.HTMLBody = strBody
		Else
			.TextBody = strBody
		End If
		.Send
	End With
	Set objMessage = Nothing
	Set Fields = Nothing
	Set objMessage = Nothing
			
	'Do error checking...
	If Err.number <> 0 then
	  SendEmail = False
	End If
	
	On Error Goto 0
End Function

Function IsSpam(str)
	IsSpam = False
	If InStr(str, "[URL") > 0 OR InStr(str, "url]") > 0 Then
		IsSpam = True
	End If
End Function

Function ReadAllTextFile(strFilename)
    Dim BinaryStream, CharSet
    
    Const adTypeText = 2
	CharSet = "UTF-8"
	
    Set BinaryStream = CreateObject("ADODB.Stream")
    
    'Specify stream type - we want To get binary data.
    BinaryStream.Type = adTypeText
    
    'Specify charset For the source text (unicode) data.
    If Len(CharSet) > 0 Then
    	BinaryStream.CharSet = CharSet
    End If
    
    'Open the stream
    BinaryStream.Open
    
    'Load the file data from disk To stream object
    BinaryStream.LoadFromFile strFilename
    
    'Open the stream And get binary data from the object
    ReadAllTextFile = BinaryStream.ReadText
End Function

Sub SaveTextFile(strFilename, strData)
    Dim BinaryStream, CharSet
	
	CharSet = "UTF-8"
	
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    
    'Create Stream object
    Set BinaryStream = CreateObject("ADODB.Stream")
    
    'Specify stream type - we want To save text/string data.
    BinaryStream.Type = adTypeText
    
    'Specify charset For the source text (unicode) data.
    If Len(CharSet) > 0 Then
    	BinaryStream.CharSet = CharSet
    End If
    
    'Open the stream And write binary data To the object
    BinaryStream.Open
    BinaryStream.WriteText strData
    
    'Save binary data To disk
    BinaryStream.SaveToFile strFilename, adSaveCreateOverWrite
End Sub

Sub AppendTextFile(strFilename, strData)
    Dim BinaryStream, CharSet
	
	CharSet = "UTF-8"
	
    Const adTypeText = 2
    Const adSaveCreateOverWrite = 2
    
    'Create Stream object
    Set BinaryStream = CreateObject("ADODB.Stream")
    
    'Specify stream type - we want To save text/string data.
    BinaryStream.Type = adTypeText
    
    'Specify charset For the source text (unicode) data.
    If Len(CharSet) > 0 Then
    	BinaryStream.CharSet = CharSet
    End If
    
    'Open the stream And write binary data To the object
    BinaryStream.Open
    BinaryStream.WriteText ReadAllTextFile(strFilename) & strData
    
    'Save binary data To disk
    BinaryStream.SaveToFile strFilename, adSaveCreateOverWrite
End Sub

Function DeleteFile(strFile)
	Dim fso
	DeleteFile = True
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(strFile) Then
		DeleteFile = fso.DeleteFile(strFile)
	End If
	Set fso = Nothing
End Function

Sub CreatePrettyURLFiles(cn, rs)
	Dim strFile, strFile2, arySites, i, strHTAccess
	
	strFile     = ""
	strHTAccess =               "RewriteEngine on" & VbCrLf
	strHTAccess = strHTAccess & "RewriteBase /" & VbCrLf
	strHTAccess = strHTAccess & "RewriteMap cocktails txt:cocktail-recipes.txt" & VbCrLf
	strHTAccess = strHTAccess & "RewriteMap products txt:products.txt" & VbCrLf
	strHTAccess = strHTAccess & "RewriteMap categories txt:categories.txt" & VbCrLf

	strHTAccess = strHTAccess & "RewriteRule ^shop/products/([^.?/]+)(\.asp) /shop/$1/ [NC,QSA,R=301]" & VbCrLf
	strHTAccess = strHTAccess & "RewriteRule ^cocktail-recipe/([^.?/]+)(\.htm) /cocktails/recipe.asp?ID=${cocktails:$1} [NC,QSA]" & VbCrLf
	strHTAccess = strHTAccess & "RewriteRule ^shooter-recipe/([^.?/]+)(\.htm) /cocktails/recipe.asp?ID=${cocktails:$1} [NC,QSA]" & VbCrLf
	strHTAccess = strHTAccess & "RewriteRule ^shop/([^.?/]+)(\.htm) /shop/viewproduct.asp?ID=${products:$1} [NC,QSA]" & VbCrLf
	strHTAccess = strHTAccess & "RewriteRule ^shop/([^.?/]+)/ /shop/viewcategory.asp?ID=${categories:$1} [NC,QSA]" & VbCrLf
	Call SaveTextFile(Server.MapPath("/.htaccess"), strHTAccess)

	rs.open "SELECT ID, name FROM cocktail ORDER BY accessed DESC", cn
	While NOT rs.EOF
		strFile  = strFile & GeneratePrettyURL(strOutDB(rs("name"))) & VbTab & rs("ID") & VbCrLf
		rs.MoveNext
	Wend
	rs.close
	Call SaveTextFile(Server.MapPath("/cocktail-recipes.txt"), strFile)
	
	strFile = ""
	rs.open "SELECT ID, name FROM DSproduct WHERE status=1 ORDER BY ID", cn
	While NOT rs.EOF
		strFile  = strFile & GeneratePrettyURL(strOutDB(rs("name"))) & VbTab & rs("ID") & VbCrLf
		rs.MoveNext
	Wend
	rs.close
	Call SaveTextFile(Server.MapPath("/products.txt"), strFile)
	
	strFile = ""
	rs.open "SELECT ID, url FROM DScategory WHERE hidden=0 ORDER BY ID", cn
	While NOT rs.EOF
		strFile  = strFile & GeneratePrettyURL(strOutDB(rs("url"))) & VbTab & rs("ID") & VbCrLf
		rs.MoveNext
	Wend
	rs.close
	Call SaveTextFile(Server.MapPath("/categories.txt"), strFile)
End Sub

Function GeneratePrettyURL(strName)
	GeneratePrettyURL = Trim(strName)
	GeneratePrettyURL = Replace(GeneratePrettyURL, " ", "-")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "-%-", "-Percent-")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "%-", "-Percent-")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "%", "-Percent")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "+-", "-")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "+", "")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "è", "e")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "&", "And")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "#", "")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "®", "")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "*", "")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "'", "")
	GeneratePrettyURL = Replace(GeneratePrettyURL, ".", "-")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "?", "-")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "/", "-")
	GeneratePrettyURL = Replace(GeneratePrettyURL, ",", "")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "|", "")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "--", "-")
	GeneratePrettyURL = Replace(GeneratePrettyURL, "--", "-")
End Function

Sub PrettyURLRedirectCocktail(cn, rs, intID, strURL)
	Dim X, strQS, strType
	
	If InStr(Request.ServerVariables("HTTP_X_REWRITE_URL"), "recipe.asp") > 0 Then
		rs.open "SELECT name, type FROM cocktail WHERE ID=" & strIntoDB(intID), cn
		If NOT rs.EOF Then
			IF Int(rs("type")) AND 1 THEN
				strType = "Cocktail"
			ELSEIF Int(rs("type")) AND 2 THEN
				strType = "Shooter"
			END IF
			strURL = "/" & strType  & "-Recipe/" & GeneratePrettyURL(strOutDB(rs("name"))) & ".htm"
		End If
		rs.close

		For Each X In Request.QueryString
			If Request(X) <> "" AND LCase(X) <> "id" Then
				strQS = strQS & X & "=" & Request(X) & "&"
			End If
		Next

		If strQS <> "" Then
			strURL = strURL & "?" & strQS
			If Right(strURL, 1) = "&" Then
				strURL = Left(strURL, Len(strURL)-1)
			End If
		End If
	End If
End Sub

Function GetURL(strUrl)
	Dim objHttp, lResolve, lConnect, lSend, lReceive, lTotal, GotResponse, intSecondsWait
	Dim blnDrinkstuff

	blnDrinkstuff = (InStr(strUrl, "drinkstuff"))
	
	Err.Clear
	On Error Resume Next
		
	lResolve 	= 5 * 1000
	lConnect 	= 5 * 1000
	lSend 		= 5 * 1000
	lReceive 	= 5 * 1000
	lTotal		= 5
	
	Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
	objHttp.open "GET", Replace(strUrl, "&amp;", "&"), False
	objHttp.send "" 
	
	intSecondsWait	= 0
	GotResponse 	= False
	Do While objHttp.readyState <> 4

		If Err.Number <> 0 Then
			Exit Do
		End If
		
		objHttp.waitForResponse 1
		intSecondsWait = intSecondsWait + 1
		
		If objHttp.readyState = 4 Then
			GotResponse = True
			Exit Do
		End If
		If intSecondsWait > lTotal Then
			GotResponse = False
			Exit Do
		End If
	Loop

	If objHttp.readyState = 4 Then
		GotResponse = True
	End If
	
	If GotResponse AND Err.Number = 0 Then
		If objHttp.status = 200 Then
			If InStr(strURL, ".jpg") > 0 OR InStr(strURL, ".gif") > 0 OR InStr(strURL, ".png") > 0 OR InStr(strURL, ".jpeg") > 0 Then
				GetURL = objHTTP.ResponseBody
			Else
				GetURL = objHTTP.ResponseText 
			End If
		End If
		
	ElseIf Err.Number <> 0 Then
		Err.Clear
	End If
	
	Set objHttp = Nothing
	On Error Goto 0

	If blnDrinkstuff Then
		GetURL = Replace(GetURL, "src=""", "src=""http://www.drinkstuff.com/")
	End If
End Function

Function stripHTML(strHTML)
'Strips the HTML tags from strHTML

  Dim objRegExp, strOutput
  Set objRegExp = New Regexp

  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  objRegExp.Pattern = "<(.|\n)+?>"

  'Replace all HTML tag matches with the empty string
  strOutput = objRegExp.Replace(strHTML, "")
  
  stripHTML = strOutput    'Return the value of strOutput

  Set objRegExp = Nothing
End Function
%>