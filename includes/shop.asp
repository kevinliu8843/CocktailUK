<%
Dim arySQLTables(24), aryIndexes(25), strDrinkstuffServer, strSQLPrefix, dblUVATRate
CONST SITE_ID = 1
strSQLPrefix = "DS"
dblUVATRate = 20

strDrinkstuffServer = "89.200.141.23"
strDrinkstuffServer = "admin.drinkstuff.com"

arySQLTables(0)  = "jointrawproduct"
arySQLTables(1)  = "jointaffiliate"
arySQLTables(2)  = "jointaffiliatelist"
arySQLTables(3)  = "jointhomepage"

arySQLTables(4)  = "dsproduct"
arySQLTables(5)  = "dsproductver"
arySQLTables(6)  = "dsrawproductver"
arySQLTables(7)  = "dsimage"
arySQLTables(8)  = "dscategory"
arySQLTables(9)  = "dsprodcat"
arySQLTables(10) = "dsprodallowdelivery"

arySQLTables(11) = "barproduct"
arySQLTables(12) = "barproductver"
arySQLTables(13) = "barrawproductver"
arySQLTables(14) = "barimage"
arySQLTables(15) = "barcategory"
arySQLTables(16) = "barprodcat"
arySQLTables(17) = "barprodallowdelivery"

arySQLTables(18) = "jointreview"
arySQLTables(19) = "jointcustomercomments"
arySQLTables(20) = "jointnonpackingdays"

arySQLTables(21) = "dsdelivery"
arySQLTables(22) = "dsdelgroup"
arySQLTables(23) = "bardelivery"
arySQLTables(24) = "bardelgroup"

aryIndexes(0) = "ALTER TABLE [dbo].[dsimage] WITH NOCHECK ADD CONSTRAINT [PK_dsimage] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(1) = "ALTER TABLE [dbo].[dsprodallowdelivery] WITH NOCHECK ADD CONSTRAINT [PK_dsprodallowdelivery] PRIMARY KEY CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(2) = "ALTER TABLE [dbo].[dsprodcat] WITH NOCHECK ADD CONSTRAINT [PK_dsprodcat] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(3) = "ALTER TABLE [dbo].[dsproduct] WITH NOCHECK ADD CONSTRAINT [PK_dsproduct] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(4) = "ALTER TABLE [dbo].[dsproductver] WITH NOCHECK ADD CONSTRAINT [PK_dsproductver] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(5) = "ALTER TABLE [dbo].[dsrawproductver] WITH NOCHECK ADD CONSTRAINT [PK_dsrawproductver] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(6) = "ALTER TABLE [dbo].[jointrawproduct] WITH NOCHECK ADD CONSTRAINT [PK_jointrawproduct] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(7) = "CREATE INDEX [IX_dsimage] ON [dbo].[dsimage]([prodID], [imagesize], [type]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(8) = "CREATE INDEX [IX_dsprodallowdelivery] ON [dbo].[dsprodallowdelivery]([prodID], [delID]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(9) = "CREATE INDEX [IX_dsprodcat] ON [dbo].[dsprodcat]([catID], [prodID], [prodorder]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(10) = "CREATE INDEX [IX_dsproduct] ON [dbo].[dsproduct]([status], [name]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(11) = "CREATE INDEX [IX_dsproductver] ON [dbo].[dsproductver]([prodID], [status], [price], [prodverorder]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(12) = "CREATE INDEX [IX_dsrawproductver] ON [dbo].[dsrawproductver]([prodverID], [rawprodID], [quantity], [rawprodorder]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(13) = "CREATE INDEX [IX_jointrawproduct] ON [dbo].[jointrawproduct]([status], [stockstatus], [stock], [preorder]) ON [PRIMARY]"

aryIndexes(14) = "ALTER TABLE [dbo].[barimage] WITH NOCHECK ADD CONSTRAINT [PK_barimage] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(15) = "ALTER TABLE [dbo].[barprodallowdelivery] WITH NOCHECK ADD CONSTRAINT [PK_barprodallowdelivery] PRIMARY KEY CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(16) = "ALTER TABLE [dbo].[barprodcat] WITH NOCHECK ADD CONSTRAINT [PK_barprodcat] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(17) = "ALTER TABLE [dbo].[barproduct] WITH NOCHECK ADD CONSTRAINT [PK_barproduct] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(18) = "ALTER TABLE [dbo].[barproductver] WITH NOCHECK ADD CONSTRAINT [PK_barproductver] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(19) = "ALTER TABLE [dbo].[barrawproductver] WITH NOCHECK ADD CONSTRAINT [PK_barrawproductver] PRIMARY KEY  CLUSTERED ([ID]) WITH  FILLFACTOR = 90  ON [PRIMARY] "
aryIndexes(20) = "CREATE INDEX [IX_barimage] ON [dbo].[barimage]([prodID], [imagesize], [type]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(21) = "CREATE INDEX [IX_barprodallowdelivery] ON [dbo].[barprodallowdelivery]([prodID], [delID]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(22) = "CREATE INDEX [IX_barprodcat] ON [dbo].[barprodcat]([catID], [prodID], [prodorder]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(23) = "CREATE INDEX [IX_barproduct] ON [dbo].[barproduct]([status], [name]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(24) = "CREATE INDEX [IX_barproductver] ON [dbo].[barproductver]([prodID], [status], [price], [prodverorder]) WITH  FILLFACTOR = 90 ON [PRIMARY]"
aryIndexes(25) = "CREATE INDEX [IX_barrawproductver] ON [dbo].[barrawproductver]([prodverID], [rawprodID], [quantity], [rawprodorder]) WITH  FILLFACTOR = 90 ON [PRIMARY]"

Sub SetAffiliateAct(strAffiliate, rs, cn, m_blnAffMode)
	Dim strProdVerList ,strPriceList
	If session("affid") <> strAffiliate Then
		'Check validity of affiliate ID
		strSQL = "SELECT Top 1 * FROM dsaffiliatelist WHERE status=1 AND affID=" & strAffiliate
		rs.open strSQL, cn
		If NOT rs.EOF Then
			'Session("showaffcategory")	= CBool(rs("showcategory"))
			Session("trueaffid")		= rs("id")
			Session("affid") 			= strAffiliate
			'Session("affSiteName") 		= rs("sitename")
			'Session("affSiteURL")  		= Trim(rs("siteurl"))
			'Session("affSiteLogoURL")  	= Trim(rs("sitelogourl"))
			'Session("affCatText")  		= Trim(rs("categorytext"))
			'Session("affPageText") 		= Trim(rs("pagetext"))
			Response.Cookies("DSaff")("ID") = strAffiliate
			Response.cookies("DSaff").Expires = #Dec 31, 2015#
			m_blnAffMode = True
			rs.close
			
			'Set up prodvers in a session var as a comma seperated string
			rs.open "SELECT prodverID, specialprice from jointaffiliate WHERE affID="&Session("trueaffid"), cn, 0, 3
			WHILE NOT rs.EOF
				If rs("specialprice") <> "" Then
					strProdVerList = strProdVerList & rs("prodverID") & ","
					strPriceList = strPriceList & rs("specialprice") & ","
				End If
				rs.MoveNext
			WEND
			rs.close
			'response.write strProdVerList
			If strProdVerList <> "" Then
				Session("affprodver") = strProdVerList
				Session("affprice")   = strPriceList
			End If
		Else
			rs.close
			Session("affprodver") = ""
			Session("affprice")   = ""
			m_blnAffMode = False
			Session("affID") = ""
		End If
	End If
End Sub

Function CheckUpdateShop()
	Dim dteUpdated, cn, rs

	Exit Function

	If Session("checkShop") Then
		Set cn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")
		cn.open strDBMod
		rs.open "SELECT dteshopupdated from dsshopupdate", cn, 0, 3
		dteUpdated = rs(0)
		rs.close
		Set rs = nothing
		Session("checkShop") = False
		If DateDiff("s", dteUpdated, Now) > 60*60*24 Then
			CheckUpdateShop = True
			cn.execute("UPDATE dsshopupdate set dteshopupdated='"&Day(Now) & "-" & MonthName(Month(Now)) & "-" & Year(Now)+10 &" 05:30:00'")
		Else
			CheckUpdateShop = False
		End If
		cn.close
		Set cn = nothing
	Else
		CheckUpdateShop = False
	End If
End Function

Function UpdateProductTables()
	CONST MAX_LEN = 1

	Dim objXmlHttpCat, rs, cn, aryRows, i, j, fso, fcuk, strData
	Dim strRow, dteStart, intTotal, intSizeDownload, dteTemp, findexes, intInserted

	Server.ScriptTimeout = 10000
	dteStart = Now()

	Set fso = Server.Createobject("Scripting.FileSystemObject")
	Set fcuk = fso.CreateTextFile(Server.MapPath("/shop/cuk_update.txt"),True)
	fcuk.writeLine("Synchronisation started: "&Now())

	Set cn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	cn.CommandTimeout = 300 ' 5 minutes for each command...
	cn.open strDBMod

	dteTemp = Now()

	On Error Resume Next

	For i=0 to UBound(arySQLTables)
		intInserted = 0
		If (Request(arySQLTables(i))="ON" AND Request("selectedtables")="true") OR Request("selectedtables")="" AND arySQLTables(i)<>"" Then
			Set objXmlHttpCat= Server.CreateObject("MSXML2.ServerXMLHTTP")
			objXmlHttpCat.open "GET", "http://"&strDrinkstuffServer&"/productfeeds/cuk/"&arySQLTables(i)&".csv" , False, "BARMANS\Lee", "Smetsy#1"
			objXmlHttpCat.send ""
			
			'response.write objXmlHttpCat.ResponseText
			
			aryRows = Split(objXmlHttpCat.ResponseText, VbCrLf)
			intSizeDownload = objXmlHttpCat.GetResponseHeader("Content-Length")
			'response.write "Updating Products...<BR>"
			strData = ""
			For j=0 To UBound(aryRows)
				If Trim(aryRows(j)) <> "" Then
					If Len(strData+aryRows(j)) > MAX_LEN Then
						If Len(strData) > 0 Then
							cn.execute(strData)
							If Err.Number = 0 Then
								intInserted = intInserted + 1
							Else
								Err.Clear
							End If
						End If
						strData = Trim(aryRows(j))
					Else
						strData = strData & Trim(aryRows(j))
					End If
				End If
			Next
			If Len(strData) > 0 Then
				cn.execute(strData)
				strData = ""
			End If
			intTotal = DateDiff("s", dteTemp, Now())
			fcuk.writeLine("Table """&arySQLTables(i)&""" ("&FormatNumber(intSizeDownload/1024, 1)&"Kb) created ("&intInserted&" rows inserted) in "&Int(intTotal/60) & " mins "& intTotal-Int(intTotal/60)*60 & " secs.")
			dteTemp = Now()
		Else
			fcuk.writeLine("Table """&arySQLTables(i)&""" not requested to be created")
		End If
	Next
	fcuk.writeLine("Product retrieval finished: "&Now())
	intTotal = DateDiff("s", dteStart, Now())
	fcuk.writeLine("Product retrieval took: "&Int(intTotal/60) & " minutes "& intTotal-Int(intTotal/60)*60 & " seconds.")

	'If Request("selectedtables")<>"true" Then
		For i=0 To UBound(aryIndexes)
			cn.execute(aryIndexes(i))
		Next
		If Err.number = 0 Then
			fcuk.writeLine("Generated indexes on data.")
		Else
			fcuk.writeLine("Error generating indexes on data - " & Err.description)
			Err.Clear
		End If
	'End If
	
	call updateShopInfo()
	
	call setupCategories(fcuk)
	
	Call CreatePrettyURLFiles(cn, rs)
	
	On Error Goto 0
	
	fcuk.writeLine("Categories updated across all sites.")
	cn.execute("UPDATE dsshopupdate set dteshopupdated='"&Day(Now) & "-" & MonthName(Month(Now)) & "-" & Year(Now) &" 03:00:00'")
	cn.close
	Set cn = nothing
	Set rs = Nothing
	Set objXmlHttpCat= Nothing
	
	fcuk.writeLine("Synchronisation finished: "&Now())
	intTotal = DateDiff("s", dteStart, Now())
	fcuk.writeLine("Synchronisation took: "&Int(intTotal/60) & " minutes "& intTotal-Int(intTotal/60)*60 & " seconds.")
	fcuk.close
	Set fcuk = nothing
	Set fso = nothing
End Function

Private Sub setupCategories(fcuk)
	'On Error Resume Next
	Dim strFontColour, strURL, f, fso, strCat, rsc, cnc, strCatOpt, strCatWap, strCatLeft
	Dim ReadAllCatFile, newCatFile, objXmlHttpCat
	Set cnc = Server.CreateObject("ADODB.Connection")
	Set rsc = Server.CreateObject("ADODB.Recordset")
	
	cnc.Open strDBMod
	
	' Turn on drink category
	cnc.execute("UPDATE DScategory SET hidden=0 WHERE ID=562")

	strSQL = "SELECT name, URL, ID, name as alt, parentID from dscategory WHERE hidden=0 AND url NOT LIKE 'admin%' ORDER by catorder"
	rsc.Open strSQL, cnc
	
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	While NOT rsc.EOF 
		If rsc("parentID") = 0 then
			strCat		= strCat 	& "<img border=""0"" src=""/images/pixel.gif"" width=""4"" height=""4""><br><A href=""/shop/products/"&Trim(rsc("URL"))&".asp"" title="""&rsc("alt")&""" class=""linksin"">"&Trim(rsc("name"))&"</A>&nbsp;<IMG border=""0"" src=""/images/side_menus/smallarrowright.gif"" width=""8"" height=""8"">&nbsp;</FONT><BR>" & VbCrLf
			strCatLeft	= strCatLeft 	& "<img border=""0"" src=""/images/pixel.gif"" width=""4"" height=""4""><br>&nbsp;<IMG border=""0"" src=""/images/side_menus/smallarrow.gif"" width=""8"" height=""8"">&nbsp;<A href=""/shop/products/"&Trim(rsc("URL"))&".asp"" title="""&rsc("alt")&""" class=""linksin"">"&Left(Trim(Capitalise(LCase(rsc("name")))), 20)&"</A></FONT><br>" & VbCrLf
		End If
		
		Set f = fso.OpenTextFile(Server.MapPath("/includes/shop/template.asp"), 1, True)
		ReadAllCatFile =  f.ReadAll
		f.close
	
		ReadAllCatFile = Replace(ReadAllCatFile, "##CAT##", rsc("ID") & "")
		
		Set newCatFile = fso.CreateTextFile(Server.MapPath("/shop/products/") & "/" & Replace(Trim(rsc("url")), " ", "_") & ".asp", True)
		newCatFile.writeline(ReadAllCatFile)
		newCatFile.close
		Set newCatFile = nothing
		rsc.movenext
	wend
	
	rsc.movefirst
	
	strCatWap = ""
	
	strCatOpt = "<"&"%strScriptName = Request.ServerVariables(""SCRIPT_NAME"")%"&">"
	strCatOpt = strCatOpt & "&nbsp;<SELECT name=""shop"" ID=""shop"" class=""shopoptioncats"" onChange=""window.location.href='/shop/' + this.options[this.selectedIndex].value"">"
	strCatOpt = strCatOpt & "<OPTION value=""default.asp"">Select a department...</OPTION>"
	While NOT rsc.EOF 
		If rsc("parentID") = 0 Then
			strCatOpt = strCatOpt & "<OPTION value=""products/"&Trim(rsc("URL"))&".asp"""
			strCatOpt = strCatOpt & "<"&"%If InStr(strScriptName,""" & Trim(rsc("URL")) & """) > 0 Then %"&"> SELECTED <"&"%End if%"&">"
			strCatOpt = strCatOpt & ">"&Trim(rsc("name"))&"</OPTION>" & VbCrLf
			
			strCatWap = strCatWap & "<a href=""/wap/shop/category.asp?ID="&rsc("ID")&""">"&Replace(Trim(rsc("name")), "&", "&amp;")&"</a><br/>"
		End If
		rsc.movenext
	Wend
	strCatOpt = strCatOpt & "</SELECT>"
	rsc.close
	
	Set f = fso.CreateTextFile(Server.MapPath("/includes/shop/categories.asp"),True)
	f.writeLine(strCat)
	f.close
	Set f = nothing
	Set f = fso.CreateTextFile(Server.MapPath("/includes/shop/categoriesleft.asp"),True)
	f.writeLine(strCatLeft)
	f.close
	Set f = nothing
	Set f = fso.CreateTextFile(Server.MapPath("/includes/shop/categoriesoption.asp"),True)
	f.writeLine(strCatOpt)
	f.close
	Set f = nothing
	Set f = fso.CreateTextFile(Server.MapPath("/wap/includes/categorieswap.asp"),True)
	f.writeLine(strCatWap)
	f.close
	Set f = nothing

	Call GenerateHomePage(cnc, rsc)
	If IsObject(fcuk) Then
		fcuk.writeline("Generated homepage.")
	End If 

	Set fso = Server.Createobject("Scripting.FileSystemObject")
	strSQL = "SELECT ID, type from dsimage"
	rsc.Open strSQL, cnc, 0, 3
	Set objGet= Server.CreateObject("MSXML2.ServerXMLHTTP")
	intImages = 0
	While NOT rsc.EOF
		strImage = rsc("ID") & "." & Trim(rsc("type"))
		If NOT FSO.FileExists(Server.MapPath("/images/shop/products/"&strImage)) Then
			objGet.open "GET", "http://www.drinkstuff.com/productimg/"&strImage, False
			objGet.send ""
			call SaveBinaryData(Server.MapPath("/images/shop/products/"&strImage), objGet.ResponseBody)
			'Response.write "Getting image " & strImage & "<BR>"
			intImages = intImages + 1 
		End If
		rsc.MoveNext
	Wend
	rsc.close
	Set objGet= Nothing

	If IsObject(fcuk) Then
		fcuk.writeLine("Retrieved "&intImages&" images")
	End If

	cnc.close
	Set rsc = Nothing
	Set cnc = Nothing
	Set fso = nothing
End Sub

Sub GenerateHomePage(cn, rs)
	Dim strHP, strText, strFooter, strTable, fso, f, aryHP, i, aryTable(), iSize, strName
	Dim strProdTitle, intTbl, strImgSrc, strUrl, strExtra, aryHPOrig, iunit, aryText()
	
on error Resume Next

	ReDim aryTable(100)
	
	intTbl = 0
	iSize  = -1
	
	rs.open "SELECT ID, name from jointhomepage WHERE parentID=0 AND site="&SITE_ID&" ORDER by contenttype, homepageorder", cn
	aryHpOrig = rs.GetRows()
	rs.close
	
	strText = ""
	strTable = "<TABLE border=0 cellpadding=6 cellspacing=0 width=""100%"">"
	strFooter = ""
	
	For iunit=0 To UBound(aryHPOrig, 2)
		rs.open "SELECT * FROM jointhomepage WHERE status=1 AND (parentID="&aryHpOrig(0,iunit)&" OR (ID="&aryHpOrig(0,iunit)&" AND parentID=0))  AND site="&SITE_ID&" ORDER by homepageorder", cn
		If NOT rs.EOF Then
			aryHP = rs.GetRows()
		Else
			ReDim aryHP(-1,-1)
		End If
		rs.close
		
		i = RandomNumber(0, UBound(aryHP,2))
		
		If UBound(aryHP, 2) >= 0 Then
			Select Case aryHP(1,i)
				Case 0
					strName   = aryHPOrig(1,iunit)
					strText   = aryHP(14,i)
					iSize = iSize + 1
					ReDim Preserve aryText(1, iSize)
					aryText(0, iSize) = strName
					aryText(1, iSize) = strText
				Case 1
					If aryHP(2, i) > 0 Then 'It is a product
						rs.open "SELECT name, "&strSQLPrefix&"image.ID as imgID, "&strSQLPrefix&"image.type from "&strSQLPrefix&"product INNER JOIN "&strSQLPrefix&"image ON "&strSQLPrefix&"product.ID="&strSQLPrefix&"image.prodID WHERE imagesize=0 AND prodID="&aryHP(2,i), cn
						If not rs.EOF Then
							strProdTitle	= strOutDB(rs("name"))
							strImgSrc 		= "/images/shop/products/" & rs("imgID") & "." & strOutDB(rs("type"))
							strUrl 			= "/shop/products/product.asp?ID="& aryHP(2,i) & "&title=" & Server.URLEncode(strOutDB(rs("name")))
							strExtra 		= "<BR/>More details <a href="""&strURL&""" target="""&aryHP(6,i)&""">here...</a>"
						Else
							strProdTitle 	= "Untitled"
							strImgSrc 		= "/img/pixel.gif"
							strUrl 			= aryHP(5,i)
							strExtra 		= ""
						End If
						rs.close
					Else
						strProdTitle 	= aryHP(4,i)
						strImgSrc 		= aryHP(3,i)
						strUrl 			= aryHP(5,i)
						strExtra 		= ""
					End If
					aryTable(intTbl) = "<P><a href="""&strUrl&""" target="""&aryHP(6,i)&"""><IMG border=""0"" src="""&strImgSrc&""" align="""&aryHP(8,i)&"""/></A>" & VbCrLf
					aryTable(intTbl) = aryTable(intTbl) & "<B>"&strProdTitle&"</B><BR/>" & VbCrLf
					aryTable(intTbl) = aryTable(intTbl) & aryHP(14,i) & strExtra & "</P>"
					intTbl = intTbl + 1
			End Select
		End If
	Next
	
	ReDim Preserve aryTable(intTbl)

	For i=0 To intTbl
		If IsEven(i) Then
			strTable = strTable & "<TR>"
		End If
		strTable = strTable & "<TD width=""50%"" valign=top>" & aryTable(i) & "</TD>"		
		If NOT IsEven(i) Then
			strTable = strTable & "</TR><TR><TD><hr size=1 color=#1F4B94 width=""100%""></TD><TD><hr size=1 color=#1F4B94 width=""100%""></TD></TR>"
		End If
	Next
	
	strTable = strTable & "</TABLE>"
	
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(Server.MapPath("/includes/shop/homepage.asp"))
	strHP =  f.ReadAll
	Set f = Nothing
	For i=0 To UBound(aryText, 2)
		strHP = Replace(strHP, "##"&aryText(0, i)&"##", aryText(1, i), 1, -1, 1)
	Next
	strHP = Replace(strHP, "##TABLE##",  strTable)
	
	'Write homepage
	Set f = fso.CreateTextFile(Server.MapPath("/shop/default.asp"),True)
	f.writeLine(strHP)
	f.close
	Set f = nothing
	Set fso = nothing
End Sub

Function IsEven(lngNum)
	' Determines whether a number is even or odd.
	IsEven = Not CBool(lngNum Mod 2)
End Function

Sub updateShopInfo()

	'On Error Resume Next
	Dim objXmlHttpCat, objXmlDoc, dblEuroToGo, dblDollarToGo, fso, f, intImages, cnc, rsc, objGet
	
	Set objXmlHttpCat= Server.CreateObject("MSXML2.ServerXMLHTTP")
	objXmlHttpCat.open "GET", "http://"&strDrinkstuffServer&"/affiliate/shopinfo.asp" , False
	objXmlHttpCat.send ""
	Set objXmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	strXML = objXmlHttpCat.responseXml.xml
	objXmlDoc.loadXML(strXML)
	Set objXmlHttpCat = nothing
	
	Set dblEuroToGo = objXmlDoc.getElementsByTagName("SHOPINFO/CURRENCIES/EURO")
	Set dblDollarToGo = objXmlDoc.getElementsByTagName("SHOPINFO/CURRENCIES/DOLLAR")
	
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	Set f = fso.CreateTextFile(Server.MapPath("/includes/shop/currency.asp"),True)
	f.writeLine("<" & "%")
	f.writeLine("Dim dblEuro, dblDollar")
	f.writeLine("dblEuro = " & dblEuroToGo.item(0).text)
	f.writeLine("dblDollar = " & dblDollarToGo.item(0).text)
	f.writeLine("%" & ">")
	f.close
	Set f = nothing
	
	Set objXmlDoc = nothing
	Set objXmlHttpCat= Nothing
End Sub

Function GetProductTablesAndUpdate()
	On Error Resume Next
	Dim objXmlHttpCat
	Set objXmlHttpCat= Server.CreateObject("MSXML2.ServerXMLHTTP")
	objXmlHttpCat.open "GET", "http://"&strDrinkstuffServer&"/productfeeds/cuk/export.asp" , False, "BARMANS\Lee", "Smetsy#1"
	objXmlHttpCat.send ""
End Function

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

Function strOutDBTextArea( ByVal strString )
	strString = strString & ""
	If NOT IsNull(strString) Then
		strString = Replace(strString, """", """""")
		strString = Replace(strString, "<BR>", VbCrLf)
		strString = Trim(strString)
	End If
	strOutDBTextArea = strString
End Function

Function RemoveLeadingChars(strChar, strRemoveFrom)
	strChar = LCase(strChar)
	While (LCase(Left(strRemoveFrom, Len(strChar))) = strChar)
		strRemoveFrom = Right(strRemoveFrom, Len(strRemoveFrom)-Len(strChar))
	Wend
	RemoveLeadingChars = strRemoveFrom
End Function

Function SaveBinaryData(FileName, ByteArray)
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write ByteArray
  
  'Save binary data To disk
  BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function

Function RandomNumber(lowerbound, upperbound)
	Randomize
	RandomNumber = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

' Examples for clarity
' =============================
' Net + VAT = Gross
' =============================
' VAT Rate:       17.5%
' Price Incl VAT: £9.99 (Gross)
' VAT:            £1.49 (VAT)
' Net:            £8.50 (Net)
' =============================
' VAT Rate:       15%
' Price Incl VAT: £9.99 (Gross)
' VAT:            £1.30 (VAT)
' Net:            £8.69 (Net)
' =============================

Function CalculateVATFromGross(dblGross, dblVatRate)
	CalculateVATFromGross = Round(dblGross - (dblGross/ (1 + (dblVatRate/100))), 2)
End Function

Function CalculateNetFromGross(dblGross, dblVatRate)
	CalculateNetFromGross = Round(dblGross/ (1 + (dblVatRate/100)), 2)
End Function

Function CalculateVATFromNet(dblNet, dblVatRate)
	CalculateVATFromNet = Round(dblNet * dblVatRate/100, 2)
End Function

Function CalculateGrossFromNet(dblNet, dblVatRate)
	CalculateGrossFromNet = Round(dblNet * (1 + (dblVatRate/100)), 2)
End Function
%>
