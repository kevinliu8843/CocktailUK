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

Private Sub SetupCategories(cnc, rsc)
	Dim strFontColour, strURL, f, fso, strCat, strCatOpt, strCatWap, strCatLeft
	Dim ReadAllCatFile, newCatFile, objXmlHttpCat
	
	' Turn on drink category
	cnc.execute("UPDATE DScategory SET hidden=0 WHERE ID=562")

	strSQL = "SELECT name, URL, ID, name as alt, parentID from dscategory WHERE hidden=0 AND url NOT LIKE 'admin%' ORDER by catorder"
	rsc.Open strSQL, cnc
	
	strCatLeft = ""
	While NOT rsc.EOF 
		If rsc("parentID") = 0 then
			strCatLeft	= strCatLeft& "<div class=""item""><A href=""/shop/"&Trim(rsc("URL"))&"/"" title="""&rsc("alt")&""" class=""linksin"">"&Left(Trim(Capitalise(LCase(rsc("name")))), 20)&"</A></div>" & VbCrLf
		End If
	wend

	Call SaveTextFile(Server.MapPath("/includes/shop/categoriesleft.asp"), strCatLeft)

	rsc.close
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
	objXmlHttpCat.open "GET", "http://"&strDrinkstuffServer&"/productfeeds/cuk/export.asp" , False, "Leetracey492", "i9t7GU|4"
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
%>
