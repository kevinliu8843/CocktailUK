<%
Function GetDrink(rs, cn, intID, aryDrink)
	GetDrink = GetActualDrink(rs, cn, intID, 1, aryDrink)
End Function

Function GetAwaitingDrink(rs, cn, intID, aryDrink)
	GetAwaitingDrink = GetActualDrink(rs, cn, intID, 0, aryDrink)
End Function

Function GetActualDrink(rs, cn, intID, intStatus, aryDrink)
	If NOT IsNumeric(intID) Then
		Exit Function
	End If

	Dim FSO, blnWAP, strDomain, objRe, arySearch(21), aryReplace(21), i

	arySearch(0) = "shot glass"
	arySearch(1) = "cocktail glass"
	arySearch(2) = "martini glass"
	arySearch(3) = "brandy balloon"
	arySearch(4) = "port glass"
	arySearch(5) = "sherry glass"
	arySearch(6) = "champagne saucer"
	arySearch(7) = "champagne flute"
	arySearch(8) = "flute"
	arySearch(9) = "lowball glass"
	arySearch(10) = "tumbler"
	arySearch(11) = "old fashioned glass"
	arySearch(12) = "highball glass"
	arySearch(13) = "tall glass"
	arySearch(14) = "wine glass"
	arySearch(15) = "shaker"
	arySearch(16) = "strainer"
	arySearch(17) = "strain"
	arySearch(18) = "boston"
	arySearch(19) = "pour"
	arySearch(20) = "blend"
	arySearch(21) = "blender"

	aryReplace(0) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(1) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(2) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(3) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(4) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(5) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(6) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(7) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(8) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(9) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(10) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(11) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(12) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(13) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(14) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(15) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(16) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(17) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(18) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(19) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(20) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	aryReplace(21) = "<a href=""/shop/products/search.asp?search=$1"">$1</a>"
	
	blnWAP = (InStr(Request.ServerVariables("SCRIPT_NAME"), "/wap") > 0)
	strDomain = Request.ServerVariables("SERVER_NAME")
	
	'Returns an array for the drink details
	'Position definitions:
	'0 Capitalised name
	'1 Description
	'2 Ingredients (inside an LI element)
	'3 Serves
	'4 Accessed
	'5 Rating
	'6 Number of users that rated this drink
	'7 Drink type ("cocktail" or "shooter")
	'8 Alcoholic type ("non-alcoholoc" OR "alcoholic")
	'9 XXX rating ("XXX rated" OR "Not XXX rated")
	'10 User's name who submitted the drink
	'11 Image location (relative to root)
	
	ReDim aryDrink(11)
	strSQL = "EXECUTE CUK_GETRECIPE @id=" & intID & ", @status=" & intStatus
	rs.Open strSQL, cn, 0, 3
	If NOT rs.EOF Then
		aryDrink(0)	= Trim( Replace( replaceStuffBack(rs("name")), ",", "" ))
		aryDrink(1)	= Capitalise(replaceStuffBack(rs("description") ))
		aryDrink(3)	= rs("serves")
		aryDrink(4)	= rs("accessed")
		aryDrink(5)	= rs("rate")
		aryDrink(6)	= rs("users")
	
		IF Int(rs("type")) AND 1 Then
			aryDrink(7) = "cocktail"
		Else
			aryDrink(7) = "shooter"
		End If
			
		IF Int(rs("type")) AND 4 Then
			aryDrink(8) = "non-alcoholic"
		Else
			aryDrink(8) = "alcoholic"
		End If
			
		If Int(rs("type")) AND 8 Then
			aryDrink(9) = "XXX rated"
		Else
			aryDrink(9) = "Not XXX rated"
		End If
	
		If rs("usr") & "" <> "" Then
			aryDrink(10) = replaceStuffBack(rs("usr"))
		Else
			aryDrink(10) = ""
		End If
	
		Set FSO = CreateObject("Scripting.FileSystemObject")
		If (FSO.FileExists(Server.MapPath("/images/cocktails/" & Replace(Server.URLEncode(Replace(Replace(aryDrink(0), "?",""),"*","")), "+", " ") & ".jpg" ))) Then
			aryDrink(11) = "<IMG src=""http://"&strDomain&"/images/cocktails/" & Replace(Server.URLEncode(Replace(Replace(aryDrink(0), "?",""),"*","")), "+", " ") & ".jpg"" ALT="""&aryDrink(0)&""" border=""0"">"
		Else
			aryDrink(11) = "<IMG src=""http://"&strDomain&"/images/cocktails/default/noimage.gif"" border=""0""><BR><IMG src=""http://"&strDomain&"/images/cocktails/default/"&aryDrink(7)&Int((6 - 1) * Rnd + 1)&".jpg"" border=""0"">"
		End If
		Set FSO = Nothing
		
		aryDrink(1) = ReplaceStuffBack(Replace(aryDrink(1), "<BR>", "<br/>", 1, -1, 1))
		If NOT blnWAP Then
			set objRE=server.createobject("VBScript.Regexp")
			objRe.Global = True
			objRe.IgnoreCase = True
			For i=0 To UBound(arySearch)
				objRe.Pattern = "\b("&arySearch(i)&")\b"
				aryDrink(1) = objRe.Replace(aryDrink(1), aryReplace(i))
			Next
			Set objRe = nothing
		End If
	Else
		GetActualDrink = False
		rs.close
		Exit Function
	End If
	rs.close
	aryDrink(2) = ReplaceStuffBack(GetRecipe(rs, cn, intID, (NOT blnWAP)))
	GetActualDrink = True
End Function

Function GetRecipe(l_rs, l_conn, intID, blnHTML)
	Dim aryRecipe, i
	aryRecipe = GetRecipeArray(l_rs, l_conn, intID)
	If IsArray(aryRecipe) Then
		GetRecipe = GetRecipe & "<UL>"
		For i=0 To UBound(aryRecipe, 2)
			GetRecipe = GetRecipe & "<LI>"
			If aryRecipe(1, i) <> "no measure" Then
				GetRecipe = GetRecipe & Trim(aryRecipe(1, i)) & "&nbsp;"
			Else
				GetRecipe = GetRecipe
			End If
			GetRecipe = GetRecipe & "<a href='/cocktails/containing.asp?ingredient=" & aryRecipe(2, i) & "' TITLE='Cocktails containing " & aryRecipe(3, i) & "''>"
			GetRecipe = GetRecipe & aryRecipe(3, i)
			GetRecipe = GetRecipe & "</a></LI>"
		Next
		GetRecipe = GetRecipe & "</UL>"
	End If
End Function

Function GetRecipeArray(l_rs, l_conn, intID)
	' Input		- a recipe ID
	' Output	- a 2 dimensional array, each row is [MeasureID, Measure, IngredientID, Ingredient]

	Dim arySplit, aryIngredientPair, i, j, strSQL
	Dim aryResult

	If IsNumeric(intID) Then
		strSQL = "EXECUTE CUK_GETRECIPEINGREDIENTS @ID="&intID
		l_rs.Open strSQL, l_conn, 0, 3
		If Not l_rs.EOF Then
			GetRecipeArray = rs.GetRows()
		End If
		l_rs.Close
	Else
		GetRecipeArray = ""
	End If
End Function

Function canIMakeIt(cn, rs, cocktailID, memID, strReturn)
	Dim bCanBeMade, strExtra
	rs.open "EXECUTE CUK_NEEDEDINGREDIENTS @c="&cocktailID&", @m="&memID, cn, 0, 3
	bCanBeMade = rs.EOF
	If bCanBeMade Then
		canIMakeIt = True 
		strReturn = "<div>You have all the ingredients you need in <A HREF=""/account/selectIngredients.asp"">your bar</A> to make this</div>"
	Else
		Do While Not rs.EOF
			If strExtra <> "" Then strExtra = strExtra & ", "
			strExtra = strExtra & "<B>" & rs("Name") &"</B>" 
			rs.MoveNext
		Loop
		canIMakeIt = False
		strReturn = "<TABLE cellpadding=2 border=0 width=""100%""><TR><TD width=""35""><IMG src=""/images/warning.gif""></TD><TD>You need <B>more</B> ingredients to make this drink : " & strExtra & "</TD></TR></TABLE>"
	End If
	rs.close
End Function

Sub writeCocktailList(strSQL, rs, cn, strTitle, strHrefType)
	Dim iPageCurrent, iPageSize, iPageCount, FSO, k, iStart, iFinish, maxPages
	Dim iWidth, iHeight, iKnt1, iKnt2, blnWAP
	blnWAP = (InStr(Request.ServerVariables("SCRIPT_NAME"), "/wap") > 0)
	' Retrieve page to show or default to 1
	If Request("page") <> "" AND IsNumeric(Request("page")) Then
		iPageCurrent = Int(Request("page"))
	Else
		iPageCurrent = 1
	End If

	iWidth = 3
	if blnWAP Then
		iWidth = 1
	End If
	iHeight = 15
	iPageSize = iWidth * iHeight

	rs.PageSize = iPageSize
	rs.CacheSize = iPageSize
	rs.CursorLocation = 3
	rs.Open strSQL, cn
	
	If NOT rs.EOF Then
		iPageCount = rs.PageCount

		' If the request page falls outside the acceptable range,
		' give them the closest match (1 or max)
		If iPageCurrent > iPageCount Then iPageCurrent = iPageCount
		If iPageCurrent < 1 Then iPageCurrent = 1

		' Move to the selected page
		rs.AbsolutePage = iPageCurrent
		Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		If Len(strTitle) >0 AND strTitle <> "" Then%>
			<h2><%=strTitle%></h2>
		<%End If%>


		<P align="center">Page <B><%= iPageCurrent %></B> of <B><%= iPageCount %></B> (<%=rs.recordCount%> recipes)</P>

		<div class="row collapse">
		  <%For iKnt2=1 To iWidth * iHeight%>
		    <div class="large-4 small-12 column">
				<%writeField FSO, rs%>
		    </div>
		  <%Next%>
		</div>
		<%Set FSO = Nothing%>

		<div class="pagination-centered">
			<ul class="pagination">
				<%
				If iPageCurrent <> 1 Then
					%>
					<li class="arrow"><a href="<%=Request.ServerVariables("URL")%>?page=<%= iPageCurrent - 1 %><%=strHrefType%>">Prev</a></li>
					<%
				End If

				k=0
				maxPages = 16
				iStart = Max(1, iPageCurrent-maxPages/2)
				iFinish = Min(iStart+maxPages-1, iPageCount)

				If (iFinish-iStart<=maxPages) Then
					iStart = iFinish - maxPages + 1
					if iStart<1 then
						iStart=1
					end if
				End If

				For k = iStart to iFinish step 1
					if iPageCurrent = k THEN%>
						<li class="current"><a class="page active" href="<%=Request.ServerVariables("URL")%>?page=<%=k%><%=Server.HTMLEncode(strHrefType)%>"><%=k%></a></li>
					<%Else%>
						<li><a class="page gradient" href="<%=Request.ServerVariables("URL")%>?page=<%=k%><%=Server.HTMLEncode(strHrefType)%>"><%=k%></a></li>
					<%End If%>
				<%
				Next

				If iPageCurrent < iPageCount Then
					%>
					<li class="arrrow"><a class="page gradient" href="<%=Request.ServerVariables("URL")%>?page=<%= iPageCurrent + 1 %><%=strHrefType%>">Next</a></li>
					<%
				End If
				%>
			</ul>
		</div>
		<%
	Else
		Response.write("<P><B>Sorry, no drinks found</B><BR><A href=""javascript:history.go(-1)"">Go back</A>")
	End If
End Sub

Sub writeField(FSO, rs)
	Dim name, fileExists, strType
	If NOT rs.EOF Then
		IF Int(rs("type")) AND 1 THEN
			strType = "Cocktail"
		ELSEIF Int(rs("type")) AND 2 THEN
			strType = "Shooter" 
		END IF
		%>
		<div class="row collapse">
			<div class="column small-1">
				<A href="/<%=strType%>-Recipe/<%=GeneratePrettyURL(replaceStuffBack(rs("name")))%>.htm"><IMG border="0" src="/images/<%=strType%>_small.gif"></A>
			</div>
			<div class="column small-11">
				<A style="padding-bottom: 3px;" href="/<%=strType%>-Recipe/<%=GeneratePrettyURL(replaceStuffBack(rs("name")))%>.htm"><%=Capitalise(replaceStuffBack(rs("name"))) %></A>
			</div>
		</div>
    	<%rs.MoveNext%>
	<%
	End If
End Sub
%>