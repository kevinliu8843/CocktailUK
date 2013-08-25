<%
Class CProduct
	Private m_intCategory
	Private m_intProduct
	Public m_strCategoryName
	Public  m_strProductName 
	Public m_strProductExtraTitle
	Private m_searchQuery
	Private m_pageSize
	Private m_numReviews
	Private m_strCatExtraTitle
	Private m_strCatExtraKeyWords
	Public  m_blnProductExists
	Private m_strKeywords
	Public  m_HardProductTitle
	Private m_AffPageText 
	Public  m_blnValidAffiliate
	Public  m_blnDisplayAffProducts
	Private m_blnAffMode
	Public  m_blnOnlyProduct
	Private m_blnWAP, m_blnMobile
	Public  m_strNoScript
	Public  m_strMetaDescription
	Private m_strBannerTarget
	Private m_strBannerURL
	Private m_strBannerImageSrc
	Public  m_dteNewProductFor
	Public  m_strCatHeader
	Private m_blnGotSubCats
	Private m_blnFlowText
	Private blnPricesInclVAT, m_intVATChargeable
	Private strCurrencySymbol, dblCurrencyFactor
	Public rs
	Public cn
	
	Public Sub Class_Initialize()
		Set cn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")
		Call Reset()
		cn.open strDB
	End Sub

	Public Sub Reset()
		m_intCategory = 0
		m_pageSize = 10
		If Session("pagesize") <> "" Then
			If IsNumeric(Session("pagesize")) then
				m_pageSize = Int(Session("pagesize"))
			End If
		End If
		m_numReviews = 5
		m_blnProductExists = False
		m_blnValidAffiliate= False
		m_strMetaDescription = ""
		m_dteNewProductFor = 45 'Days
		blnPricesInclVAT = True
		m_intVATChargeable = True
		dblCurrencyFactor = 1
	End Sub

	Public Sub Class_Terminate()
		cn.close
		Set cn = nothing
		Set rs = nothing
	End Sub
	
	Public Sub SetPageSize(intSize)
		If IsNumeric(intSize) Then
			m_pageSize	= intSize
		End If
	End Sub

	Public Sub SetCategory(intCategory)
		If IsNumeric(intCategory) Then
			m_intCategory	= intCategory
		End If
	End Sub

	Public Sub SetProductID(intProduct)
		If IsNumeric(intProduct) Then
			m_intProduct	= intProduct
		End If
	End Sub
	
	Public Function GetCategoryKeywords()
		GetCategoryKeywords= m_strCatExtraKeyWords
	End Function

	Public Function DisplayTitle()
		If m_strCategoryName = "" AND m_intCategory > 0 Then
			call GetCategoryName()
		ElseIf m_intProduct > 0 Then
			call GetProductName()
		End if
		
		If m_intCategory > 0 Then
			DisplayTitle = m_strCategoryName 
		ElseIf m_intProduct > 0 Then
			DisplayTitle = m_strProductName & m_strProductExtraTitle
		ElseIf m_searchQuery <>  "" then
			DisplayTitle = "Search for " & Replace(m_searchQuery, "''", "'")
		ElseIf Session("affid") <> "" Then
			DisplayTitle = Session("affCatText")
		End If
	End function
	
	Public Function DisplayTopTitle()
		If m_strCategoryName = "" AND m_intCategory > 0 Then
			call GetCategoryName()
		ElseIf m_intProduct > 0 then
			call GetProductName()
		End if
		
		If m_intCategory > 0 Then
			DisplayTopTitle = "Buy " & m_strCategoryName & " | Cocktail : UK Bar Equipment Shop"
		ElseIf m_intProduct > 0 Then
			DisplayTopTitle = "Buy " & m_strProductName & " | Cocktail : UK Bar Equipment Shop"
		ElseIf m_searchQuery <>  "" then
			'DisplayTopTitle = "Drinkstuff.com - search for " & Replace(m_searchQuery, "''", "'")
		ElseIf Session("affid") <> "" Then
			'DisplayTopTitle = "Drinkstuff.com - " & Session("affCatText")
		End If
	End Function
	
	Public Sub GetCategoryName()
		m_strCategoryName = ""
		If m_intCategory > 0 Then
			strSQL = "SELECT extratitle, extrakeywords, noscript, bannertarget, bannerURL, bannertype FROM dscategoryactual WHERE catID=" & m_intCategory
			rs.open strSQL, cn, 0, 3
			If NOT rs.EOF Then
				m_strCatExtraTitle		= strOutDB(rs("extratitle"))
				m_strCatExtraKeyWords	= strOutDB(rs("extrakeywords"))
				m_strNoScript			= strOutDB(rs("noscript"))
				m_strBannerTarget 		= strOutDB(rs("bannertarget"))
				m_strBannerURL 			= strOutDB(rs("bannerURL"))
				m_strBannerImageSrc 	= "/images/shop/banners/" & m_intCategory & "." & strOutDB(rs("bannertype"))
			End If
			rs.close

			strSQL = "SELECT name, extratitle, extrakeywords, header FROM dscategory WHERE ID=" & m_intCategory
			rs.open strSQL, cn, 1, 3
			If NOT rs.EOF Then
				m_strCategoryName 	= strOutDB(rs("name"))
				'm_strCatHeader	= rs("header")
				If m_strCatHeader <> "" then
					m_strCatHeader = Replace(m_strCatHeader, "src=""", "src=""http://www.drinkstuff.com")
				End If
				If m_strCatExtraTitle = "" Then
					m_strCatExtraTitle	= strOutDB(rs("extratitle"))
				End If
				If m_strCatExtraKeyWords = "" Then
					m_strCatExtraKeyWords	= strOutDB(rs("extrakeywords"))
				End If
			End If
			rs.close
		End If
	End Sub
	
	Public Sub GetProductName()
		m_strProductName = ""
		m_strProductExtraTitle = ""
		If m_intProduct > 0 then
			strSQL = "SELECT name, keywords, subtext, title FROM dsproduct INNER JOIN dsproductver ON dsproduct.ID=dsproductver.prodID WHERE dsproduct.status=1 AND dsproductver.status=1 AND dsproduct.ID=" & Min(CDbl(m_intProduct),99999)
			rs.open strSQL, cn
			WHILE NOT rs.EOF
				m_strProductName = strOutDB(rs("name"))
				m_strKeywords	 = strOutDB(rs("keywords"))
				m_HardProductTitle = strOutDB(rs("title"))
				If strOutDB(rs("subtext")) <> "" Then
					If m_strProductExtraTitle = "" Then
						m_strProductExtraTitle = m_strProductExtraTitle & " ("
					End If
					m_strProductExtraTitle = m_strProductExtraTitle & strOutDB(rs("subtext")) & "," 
				End If
				rs.movenext
			Wend
			If m_strProductExtraTitle <> "" Then
				m_strProductExtraTitle = m_strProductExtraTitle & ")"
			End If
			rs.close
		End If
	End Sub
	
	Public Sub GetKeywords(strKeywords, strKeywordsOut)
		If m_strKeywords <> "" then
			strKeywordsOut = m_strKeywords
		Else
			strKeywordsOut = strKeywords
		End If	
	End Sub

	Public Sub DisplayNoResults(intType)
		Select Case intType
			Case 0
				If NOT m_blnGotSubCats Then
					Response.write("There are no products to display...")
				End If
			Case 1
				Response.write("Sorry, there appears to be no product with the code specified.")
			Case 2
				Response.write("Sorry, we were unable to find anything for " & m_searchQuery & "<br/>Please try relaxing your search term.")
			End Select
	End Sub
	
	Private Function ChangeMacros(strTextIn)
		Dim strTextOut
		strTextOut = strTextIn
		strTextOut = Replace(strTextOut, "www.drinkstuff.com/products", "www.cocktail.uk.com/shop/products")
		strTextOut = Replace(strTextOut, "src=""/", "src=""http://www.drinkstuff.com/")
		strTextOut = Replace(strTextOut, "src=/", "src=http://www.drinkstuff.com/")
		strTextOut = Replace(strTextOut, "src='/", "src='http://www.drinkstuff.com/")
		strTextOut = Replace(strTextOut, "##PHONE##", "0875 428 0958")
		ChangeMacros = strTextOut
	End Function
	
	Public Sub DisplayProducts()
		Dim aryRows, i, intNextID, blnDrawTopTable , blnDrawBottomTable, blnDoOption, intStock
		Dim iPageCurrent, iPageCount, iRecordsShown, iPageSize, intCurrentProduct, intTotalProducts
		Dim dblPrice, strAffCaption, strQuery, blnCollectionOnly, blnOutOfStock, strSdesc, blnSubCats
		Dim dteDueIn, blnPreorder
		
		If m_intCategory > 0 Then	
			strSQL = "DS_DISPLAYPRODUCTS @catID=" & m_intCategory
			
			If Request.QueryString("page") = "" OR NOT IsNumeric(Request.QueryString("page")) Then
				iPageCurrent = 1
			ElseIf IsNumeric(Request.QueryString("page")) Then
				iPageCurrent = Int(Min(CDbl(Request.QueryString("page")),99999))
			End If
			
			If iPageCurrent < 1 Then iPageCurrent = 1
			
			rs.open strSQL, cn
			' Retrieve page to show or default to 1
						
			If NOT rs.EOF Then
				aryRows = rs.GetRows()
			Else
				Redim aryRows(0,0)
			End If
			rs.close

			rs.open "DS_GETPARENTCAT @catID="&m_intCategory, cn
			If NOT rs.EOF Then
				%>
				<P style="margin-bottom: 1em"><b><a href="/shop/<%=GeneratePrettyURL(rs("url"))%>/" style="text-decoration: none;">
				&laquo; Back To <span style="text-decoration: underline;"><%=strOutDB(rs("name"))%></span></a></b></p>
				<%
			End If
			rs.close
			
			strSQL = "DS_GETSUBCATS @catID="&m_intCategory
			rs.open strSQL, cn
			If NOT rs.EOF Then
				m_blnGotSubCats = True
				i=0
				%><div class="row collapse"><%
				WHILE NOT rs.EOF
					i=i+1
					%><div class="large-4 column small-6" style="padding-bottom:25px;"><A href="/shop/<%=GeneratePrettyURL(strOutDB(rs("url")))%>/" style="font-weight: 400; font-size: 130%">
					<%Call GetSubCategoryImage(rs("ID"))%><%=strOutDB(rs("name"))%>&nbsp;&raquo;</A></div>
					<%rs.movenext
				Wend
				%>
				</div>
				<%
			End If
			rs.close
			
			If m_blnGotSubCats AND UBound(aryRows,1) > 0 Then
				%><H3>Featured products</H3><%
			End If

			intCurrentProduct = 0

			If m_strCatHeader <> "" Then
			   %><p><%=m_strCatHeader%></p><%
			End if

			For i=0 to UBound(aryRows,2)
				If intCurrentProduct >= (iPageCurrent-1)*m_pageSize AND intCurrentProduct < iPageCurrent*m_pageSize Then
					%>
					<div class="row collapse">
					    <div class="small-3 column">
						    <A href="/shop/<%=GeneratePrettyURL(strOutDB(aryRows(1,i)))%>.htm" title="<%=strOutDB(aryRows(1,i))%>"><IMG style="max-width: 90%; width: 90%;" border="0" src="http://www.drinkstuff.com/productimg/<%=strOutDB(aryRows(3,i))%>.<%=strOutDB(aryRows(2,i))%>"></A>
						</div>
						<div class="small-9 column">
							<div style="padding-right: 5%">
						        <h4><A href="/shop/<%=GeneratePrettyURL(strOutDB(aryRows(1,i)))%>.htm"><%=strOutDB(aryRows(1,i))%></a><span style="color: black"> - 
						        	<%If aryRows(4,i) <> aryRows(5,i) then%>From <%End If%>&pound;<%=FormatNumber(aryRows(4,i), 2)%></span></h4>
						        <p><%=ChangeMacros(strOutDB(aryRows(6,i)))%> <a href="/shop/<%=GeneratePrettyURL(strOutDB(aryRows(1,i)))%>.htm">[more...]</a></p>
						    </div>
					    </div>
					</div>
					<hr>
					<%
				End If
				
				intCurrentProduct = intCurrentProduct + 1
			Next

			iPageCount = Int(((UBound(aryRows, 2)-1) / m_pageSize)+1)
			If iPageCount > 1 then
				%>
				<div class="pagination-centered">
					<ul class="pagination">
					<%
					If iPageCurrent <> 1 Then
						%>
						<li class="arrow"><a href="?page=<%= iPageCurrent - 1 %>">Prev</a></li>
						<%
					End If

					For I = 1 To iPageCount
						If I = iPageCurrent Then%>
							<li class="current"><%=I%></li>
						<%Else%>
							<li><a href="?page=<%= I %>"><%=I%></a></li>
							<%
						End If
					Next
				
					If iPageCurrent < iPageCount Then
						%>
						<li class="arrrow"><a class="page gradient" href="?page=<%= iPageCurrent + 1 %>">Next</a></li>
						<%
					End If
					%>
					</ul>
				</div>
				<%
			End if
		End If
	End Sub
	
	Public Sub DisplayProduct()
		Dim aryRows, i, j, k, blnDoOption, intNumImages, dblPrice, strAffCaption, strLdesc
		Dim strBackURL, blnCollectionOnly, blnOutOfStock , dblMaxPrice, dblMinPrice, intRewardPoints
		Dim intStock, dteDueIn, blnPreOrder, aryExtraNormalImages, intMinStock, intPreOrder, intMaxStock

		If m_intProduct > 0 then
			strSQL = "DS_DISPLAYPRODUCT @intProduct=" & Min(CDbl(m_intProduct),99999) & ", @imgSize=1"
			
			rs.open strSQL, cn
			If NOT rs.EOF Then
				aryRows = rs.GetRows()
			Else
				Redim aryRows(0,0)
			End If
			rs.close

			If UBound(aryRows,2) >= 0 then
				%>
				<P align="center"><img src="http://www.drinkstuff.com/productimg/<%=aryRows(4,0)%>.<%=Trim(aryRows(5,0))%>"></p>
				<a href="/shop/delivery.asp" style="float: right;">
				<img alt="Free UK Delivery" src="../../images/shop/Free-Delivery.gif" width="135" height="122" class="style1"></a>
				<p>
				<%
				Response.Write ChangeMacros(strOutDB(aryRows(6,0)))
				Response.Write ChangeMacros(strOutDB(aryRows(7,0)))
				%>
				</p>
				<p align="center"><img src="/img/"></p>
				<%
			End If
		End If
	End Sub
	
	Private Sub GetSubCategoryImage(img)
		Dim objGet

		Set objGet= Server.CreateObject("Msxml2.ServerXMLHTTP")
		objGet.open "GET", "http://www.drinkstuff.com/img/categories/" & img & ".jpg", False
		objGet.setTimeouts 5000, 5000, 5000, 5000
		objGet.send ""
		If (objGet.Status = 200 AND objGet.getResponseHeader("Content-Type") = "image/jpeg") Then 
			response.write("<IMG src=""http://www.drinkstuff.com/img/categories/"&img&".jpg"" border=""0""><BR>")
			Set objGet = nothing
			Exit Sub
		End If
		Set objGet= Nothing

		Set objGet= Server.CreateObject("Msxml2.ServerXMLHTTP")
		objGet.open "GET", "http://www.drinkstuff.com/img/categories/" & img & ".gif", False
		objGet.setTimeouts 5000, 5000, 5000, 5000
		objGet.send ""
		If (objGet.Status = 200 AND objGet.getResponseHeader("Content-Type") = "image/gif") Then 
			response.write("<IMG src=""http://www.drinkstuff.com/img/categories/"&img&".gif"" border=""0""><BR>")
			Set objGet = nothing
			Exit Sub
		End If
		Set objGet= Nothing
	End Sub
End Class
%>