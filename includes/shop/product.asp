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
		'Session("affid") = 10724
		'call SetAffiliate(10724)
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
		m_blnDisplayAffProducts = False
		m_blnOnlyProduct = False
		m_blnAffMode = (Session("affid") <> "")
		m_blnWAP = False
		m_blnMobile = False
		m_strMetaDescription = ""
		m_dteNewProductFor = 45 'Days
		blnPricesInclVAT = True
		m_intVATChargeable = True
		strCurrencySymbol = ""
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
	
	Public Sub SetNumReviews(intNum)
		If IsNumeric(intNum) Then
			m_numReviews = intNum
		End If
	End Sub
	
	Public Sub SetAffiliate(strAffiliate)
		Call SetAffiliateAct(strAffiliate, rs, cn, m_blnAffMode)
		m_blnValidAffiliate = m_blnAffMode
		m_blnDisplayAffProducts = Session("showaffcategory")
	End Sub

	Public Function IsAffOffer(prodverID, dblPrice)
		Dim aryProdVer, aryPrice, i
		aryProdVer = Split(Session("affprodver"), ",")
		IsAffOffer = False
		For i=0 To UBound(aryProdVer)
			If Trim(CStr(aryProdVer(i))) = Trim(CStr(prodverID)) Then
				IsAffOffer = True
				aryPrice = Split(Session("affprice"), ",")
				dblPrice = aryPrice(i)
				Exit For
			End If
		Next
	End Function

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
			
			If m_blnWAP OR m_blnMobile Then
				DisplayTitle = m_strProductName 
			End If
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

	Public Function GetNoScript()
		If m_strNoScript = "" Then
			m_strNoScript = "<H1>"&m_strProductName&"</H1>"
		End If
		GetNoScript = m_strNoScript
	End Function
	
	Public Sub GetBanner(strBannerTarget, strBannerURL, strBannerImageSrc)
		strBannerTarget 	= m_strBannerTarget
		strBannerURL 		= m_strBannerURL
		strBannerImageSrc 	= m_strBannerImageSrc
	End Sub
	
	Public Sub DisplayReviews() 
		Dim i, iMax
		If m_intProduct > 0 Then
			strSQL = "SELECT name, review, dte FROM jointreview WHERE site=1 AND status=1 AND prodID=" & m_intProduct
			strSQL = strSQL & " ORDER by dte DESC"
			rs.open strSQL, cn
			iMax = rs.recordcount
			If NOT rs.EOF then
				response.write ("<p>")
				For i = 1 To Min(iMax, m_numReviews)
					response.write (strOutDB(rs("review")) & "<br/>&nbsp;&nbsp;&nbsp;<b><FONT size=1>" & strOutDB(rs("name")) & " </FONT><FONT size=1 color=gray>" & MonthName(Month(rs("dte")),True) & " " & Year(rs("dte")) & "</FONT></b><br/>")
					rs.movenext
				Next
				If iMax > m_numReviews Then
					Response.write("<p align=left>There are more reviews. <A href="""&Request.ServerVariables("SCRIPT_NAME")&"?" & Request.querystring & "&reviews=all"">Read them all...</A>")
				End If
			Else
				Response.write("<p align=left>There are currently no reviews for this product, be the first.")
			End If
			rs.close
		End If
	End Sub
	
	Public Sub SetSearchQuery(strQuery)
		m_searchQuery = strQuery
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
	
	Public Sub SetOnlyProduct()
		m_blnOnlyProduct = True
	End Sub
	
	Public Sub SetWAP()
		m_blnWAP = True
	End Sub
	
	Public Sub SetMobile()
		m_blnMobile = True
	End Sub
	
	Private Function ChangeMacros(strTextIn)
		Dim strTextOut
		strTextOut = strTextIn
		If InStr(strTextOut, "##FREEUKDEL##") > 0 Then
			strTextOut = Replace(strTextOut, "##FREEUKDEL##", GetFreeUkDel())
		End If
		strTextOut = Replace(strTextOut, "www.drinkstuff.com/products", "www.cocktail.uk.com/shop/products")
		'strTextOut = Replace(strTextOut, "www.drinkstuff.com", "www.cocktail.uk.com")
		strTextOut = Replace(strTextOut, "src=""/", "src=""http://www.drinkstuff.com/")
		strTextOut = Replace(strTextOut, "src=/", "src=http://www.drinkstuff.com/")
		strTextOut = Replace(strTextOut, "src='/", "src='http://www.drinkstuff.com/")
		strTextOut = Replace(strTextOut, "##PHONE##", "0875 428 0958")
		ChangeMacros = strTextOut
	End Function
	
	Private Function GetFreeUkDel()
		rs.open "SELECT min_spend_fd FROM dsdelivery WHERE defaultzone=1 AND status=1", cn
		If NOT rs.EOF Then
			GetFreeUkDel = rs(0)
		Else
			GetFreeUkDel = 35
		End If
		rs.close
	End Function
	
	Public Sub DisplayProducts()
		Dim aryRows, i, intNextID, blnDrawTopTable , blnDrawBottomTable, blnDoOption, intStock
		Dim iPageCurrent, iPageCount, iRecordsShown, iPageSize, intCurrentProduct, intTotalProducts
		Dim dblPrice, strAffCaption, strQuery, blnCollectionOnly, blnOutOfStock, strSdesc, blnSubCats
		Dim dteDueIn, blnPreorder
		
		If m_intCategory > 0 OR m_searchQuery <> "" OR m_blnDisplayAffProducts OR m_blnOnlyProduct then	
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
				<P style="margin-bottom: 1em"><b><a href="<%=strOutDB(rs("url"))%>.asp" style="text-decoration: none;">
				&laquo; Back To <span style="text-decoration: underline;"><%=strOutDB(rs("name"))%></span></a></b></p>
				<%
			End If
			rs.close
			
			Call DisplaySubCategories()
			
			If m_blnGotSubCats AND UBound(aryRows,1) > 0 Then
				%><H3>Featured products</H3><%
			End If

			If UBound(aryRows,1) <= 0 Then
				Call DisplayNoResults(0)
				Exit Sub
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

			' Show "previous" and "next" page links which pass the page to view
			' and any parameters needed to rebuild the query.  You could just as
			' easily use a form but you'll need to change the lines that read
			' the info back in at the top of the script.

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
			strSQL = "DS_DISPLAYPRODUCT @intProduct=" & Min(CDbl(m_intProduct),99999) 
			If NOT m_blnWAP AND NOT m_blnMobile Then
				strSQL = strSQL & ", @imgSize=1"
			Else
				strSQL = strSQL & ", @imgSize=0"
			End If
			
			rs.open strSQL, cn
			If NOT rs.EOF Then
				aryRows = rs.GetRows()
			Else
				Redim aryRows(0,0)
			End If
			rs.close

			'Get normal images
			rs.open "SELECT ID, type FROM DSimage WHERE prodID="&Min(CDbl(m_intProduct),99999)&" and imagesize=1 order by imageorder, ID", cn
			If NOT rs.EOF Then
				aryExtraNormalImages = rs.GetRows()
			Else
				ReDim aryExtraNormalImages(-1, -1)
			End If
			rs.close
			
			If UBound(aryRows,1) <= 0 Then
				Call DisplayNoResults(1)
				Exit Sub
			End If
			
			blnDoOption = (UBound(aryRows,2) > 0)

			If UBound(aryRows,1) > 0 then
				m_blnProductExists = True
				i=0
				
				strSQL = "SELECT count(*) from dsimage where prodID="&m_intProduct&" AND imagesize=2"
				rs.open strSQL, cn
				intNumImages = rs(0)
				rs.close
				
				If Request("previous") = "" Then
					strBackURL = Request.ServerVariables("HTTP_REFERER")
					if strBackURL = "" OR NOT (InStr(strBackURL, "cocktail") > 0 OR InStr(strBackURL, "cocktail") > 0 OR InStr(strBackURL, "cocktail") > 0 OR InStr(strBackURL, "localhost") > 0 OR InStr(strBackURL, "lee") > 0) Then
						strBackURL = "/default.asp"
					End If
				Else
					strBackURL = Request("previous")
					If InStr(strBackURL, "added=true") > 0 Then
						strBackURL = Replace(strBackURL, "added=true", "added=")
					End If
				End If
				blnCollectionOnly = IsCollectionOnly(cn, rs, Min(CDbl(m_intProduct),99999))
				
				'Find out max price
				dblMaxPrice = 0
				For k=0 To UBound(aryRows,2)
					dblMaxPrice = Max(aryRows(5,k), dblMaxPrice)
				Next
				intRewardPoints = Int(dblMaxPrice/20)
				If NOT m_blnWAP AND NOT m_blnMobile Then
				%>
				
				<%If UBound(aryExtraNormalImages, 2) > 0 Then%>
					<script type="text/javascript" src="http://yui.yahooapis.com/2.5.0/build/yahoo-dom-event/yahoo-dom-event.js"></script> 
					<script type="text/javascript" src="http://yui.yahooapis.com/2.5.0/build/utilities/utilities.js"></script>
					<script type="text/javascript" src="http://yui.yahooapis.com/2.5.0/build/dragdrop/dragdrop-min.js"></script>
					<script type="text/javascript" src="http://yui.yahooapis.com/2.5.0/build/container/container_core-min.js"></script>
					<script type="text/javascript" src="/script/carousel.js"></script>
					
					<link href="/style/carousel.css" rel="stylesheet" type="text/css">
					<style type="text/css">
					.carousel-component { 
						background: #e1e3f2; 
						padding:4px;
						margin:0px;
						width:140px; /* seems to be needed for safari */
						border: 1px solid #cccccc;
					}
					
					.carousel-component .carousel-list li { 
						margin:0px;
						padding:4px;
						width:130px;
					}
					
					/* Applies only to vertical carousels */
					.carousel-component .carousel-vertical li { 
						margin-bottom:0px;
						height:128px;
					}
					
					.carousel-component .carousel-list li a { 
						display:block;
						border:1px solid #e2edfa;
						outline:none; 
					}
					
					.carousel-component .carousel-list li a:hover { 
						border: 1px solid #aaaaaa;
					}
					
					.carousel-component .carousel-list li img { 
						display:block; 
					}
														
					#up-arrow { 
						cursor:pointer; 
						margin-left: 55px;
					}
					
					#down-arrow { 
						cursor:pointer; 
						margin-left: 55px;
					}
					</style>
					<script type="text/javascript">
					
					/**
					 * Custom button state handler for enabling/disabling button state. 
					 * Called when the carousel has determined that the previous button
					 * state should be changed.
					 * Specified to the carousel as the configuration
					 * parameter: prevButtonStateHandler
					 **/
					var handlePrevButtonState = function(type, args) {
					
						var enabling = args[0];
						var leftImage = args[1];
						if(enabling) {
							leftImage.src = "/images/template/up-enabled.gif";	
						} else {
							leftImage.src = "/images/template/up-disabled.gif";	
						}
						
					};
					
					/**
					 * Custom button state handler for enabling/disabling button state. 
					 * Called when the carousel has determined that the next button
					 * state should be changed.
					 * Specified to the carousel as the configuration
					 * parameter: nextButtonStateHandler
					 **/
					var handleNextButtonState = function(type, args) {
					
						var enabling = args[0];
						var rightImage = args[1];
						
						if(enabling) {
							rightImage.src = "/images/template/down-enabled.gif";
						} else {
							rightImage.src = "/images/template/down-disabled.gif";
						}
						
					};
					
					/**
					 * You must create the carousel after the page is loaded since it is
					 * dependent on an HTML element (in this case 'mycarousel'.) See the
					 * HTML code below.
					 **/
					var carousel; // for ease of debugging; globals generally not a good idea
					var pageLoad = function() 
					{
						carousel = new YAHOO.extension.Carousel("mycarousel", 
							{
								itemWidth:				120,
								itemHeight:				120,						
								numVisible:				2, 
								animationSpeed:			0.15,
								scrollInc:				2,
								revealAmount:			0, 
								firstVisible: 			1, 
								prevElement:			"up-arrow",
								nextElement:			"down-arrow",
								size:					<%=UBound(aryExtraNormalImages, 2)+1%>, 
								orientation:			"vertical", 
								wrap:					false,
								prevButtonStateHandler:	handlePrevButtonState,
								nextButtonStateHandler:	handleNextButtonState,
								loadInitHandler:   		loadInitialItems
							}
						); 
					};
					
					var loadInitialItems = function(type, args)
					{
						var start = args[0];
						var last = args[1]; 
						var c = this;
						c.show();
					}
					
					YAHOO.util.Event.addListener(window, 'load', pageLoad);
					
					</script>
				<%End If%>
				
				<style>
				.style1 {
					border-style: solid;
					border-width: 0;
					margin-left: 10px;
				}
				</style>
				
				<script>
				function changeimage(newsrc){
					img = document.getElementById('productimage')
					if (document.all){
					  img.style.filter="blendTrans(duration=0.3)" 
					  img.filters.blendTrans.Apply()
					}
					img.src = newsrc
					if (document.all){
					  img.filters.blendTrans.Play()
					}
				}
				</script>
	          <p align="left" style="margin-left: 5px; margin-right: 5px"> 
				<img border="0" src="/images/shop/less.gif" align="middle" hspace="3"><b><A HREF="javascript:history.go(-1)"><font color="#636388" size="2">Back</font></A></b></p>
	
	          <p align="left" style="margin-left: 5px; margin-right: 5px">  
	          	<IMG border="0" ID="productimage" src="http://www.drinkstuff.com/productimg/<%=strOutDB(aryExtraNormalImages(0,0))%>.<%=strOutDB(aryExtraNormalImages(1,0))%>" alt="<%=strOutDB(aryRows(1,i))%>" align="left">
	          </p>

			<%If UBound(aryExtraNormalImages, 2) > 0 Then%> 
				<div style="float: right; margin-right: 15px; margin-top: -20px; line-height: 1; ">
					<div style="margin:0px;"> 
					    <img id="up-arrow" class="left-button-image" src="/images/template/up-enabled.gif" alt="Previous Button">
					</div>
					<div id="mycarousel" class="carousel-component">
					  <div class="carousel-clip-region">
					    <ul class="carousel-list">
					  	<%For k=0 To UBound(aryExtraNormalImages, 2)%>
					  		<li><a href="javascript:void(0)" title="Click to view this image..." onclick="changeimage('http://www.drinkstuff.com/productimg/<%=aryExtraNormalImages(0,k)%>.<%=aryExtraNormalImages(1,k)%>')"><img border="0" src="http://www.drinkstuff.com/productimg/<%=aryExtraNormalImages(0,k)%>.<%=aryExtraNormalImages(1,k)%>" width="120" height="120"></a></li>
					  	<%Next%>
					    </ul>
					  </div>
					</div>
					<div style="margin: 0px; margin-top: 2px; ">
					    <img id="down-arrow" class="right-button-image" src="/images/template/down-enabled.gif" alt="Next Button">
					</div>
				</div>
			<%End If%>

			<%If intNumImages > 0 Then%>
				<P align="center"><A HREF="#" onClick="window.open('http://www.drinkstuff.com/products/images.asp?ID=<%=m_intProduct%>&logo=http://www.cocktail.uk.com/images/template/cuk_logo_banner.gif','image','width=600, height=500, toolbar=0, menubar=0, status=0, scrollbars=1, resizable=1')"><B><FONT size="1">
				<IMG border="0" src="/images/buttons/zoom.gif"></FONT></B></A>
			<%End If%> 
	
			<%If Request("added") <> "" Then%><p align="center"><b><U><FONT color="#FF0000">Product added to your basket</FONT></U></b></p><%End If%>
	
			<% 
			dblMinPrice	= 999999
			dblMaxPrice	= 0
			intMinStock	= 999999
			intMaxStock	= 0
			intPreOrder	= 1
			blnPreorder	= False
			For j=0 To UBound(aryRows, 2)
				Call GetProductVerPrice(aryRows, j, False, dblPrice, strAffCaption)
				dblMinPrice = Min(dblMinPrice, dblPrice)
				dblMaxPrice = Max(dblMaxPrice, dblPrice)
				intMinStock = Min(intMinStock, aryRows(11, j))
				intMaxStock = Max(intMaxStock, aryRows(11, j))
				If NOT aryRows(13, j) Then
					intPreOrder	= 0
				End If
				If aryRows(11, j) <= 0 AND aryRows(13, j) Then
					blnPreorder = True
				End If
			Next
			intMinStock = 1
			intMaxStock = 100
			aryRows(11, j) = 100
			%>
		      <FORM method="POST" action="/shop/basket.asp" name="addprod" style="clear: both">
				<div id="product-versions">
					<div id="pv-title">
					<%If intMaxStock > 0 Then%>
						<img alt="In Stock" src="../../images/shop/InStock.png" width="25" height="25" align="absmiddle"> In Stock &amp; Available For Immediate Dispatch
					<%ElseIf intMinStock > 0 Then%>
						<img alt="In Stock" src="../../images/shop/InStock.png" width="25" height="25" align="absmiddle"> In Stock &amp; Available For Immediate Dispatch
					<%ElseIf intPreOrder > 0 Then%>
						<img alt="Available To Pre-Order" src="../../images/shop/InStock.png" width="25" height="25" align="absmiddle"> Available To Pre-Order
					<%ElseIf intMaxStock <= 0 Then%>
						<img alt="Out Of Stock" src="../../images/shop/OutOfStock.png" width="25" height="25" align="absmiddle"> Currently Out Of Stock
					<%End If%>
					</div>
					<div id="product-version-container">
						<table style="width: 95%" cellspacing="0" cellpadding="0" class="product-version">
						<%For j=0 To UBound(aryRows,2)%>
							<%Call GetProductVerPrice(aryRows, j, False, dblPrice, strAffCaption)%>
							<tr>
								<td><%If strOutDB(aryRows(3, j)) <> "" Then%><%=strOutDB(aryRows(3, j))%><%Else%><%=strOutDB(aryRows(1, j))%><%End If%></td>
								<td align="left" nowrap="nowrap" width="80">
								<%If aryRows(11, j) > 0 Then%>
									<img alt="In Stock" src="../../images/shop/InStockSmall.png" width="14" height="14" align="absmiddle"> In Stock
								<%ElseIf aryRows(13, j) Then%>
									<img alt="Available to Pre-Order" src="../../images/shop/InStockSmall.png" width="14" height="14" align="absmiddle"> Available to Pre-Order
								<%Else%>
									<img alt="Out Of Stock" src="../../images/shop/OutOfStockSmall.png" width="14" height="14" align="absmiddle"> Out Of Stock
								<%End If%>
								</td>
								<td width="70" align="right"><span class="price">&pound;<%=FormatNumber(dblPrice, 2)%></span></td>
								<td width="55" align="right"><select name="quantity<%=aryRows(4, j)%>"> 
								<option value="0">Qty</option>
								<%For k=1 To 20%>
									<option value="<%=k%>" <%If UBound(aryRows,2)=0 AND k=1 Then%>selected<%End If%>><%=k%></option>
								<%Next%>
								</select></td>
							</tr>
						<%Next%>
						</table>
						<div style="text-align: right; margin-top: 5px;"><a href="http://www.awin1.com/awclick.php?gid=73406&mid=8&awinaffid=176043&linkid=101500&clickref=&p=<%=Server.URLEncode("http://www.drinkstuff.com/products/affiliate.asp?affID=987654321&prodID="&m_intProduct)%>"><img name="Image1" src="../../images/template/addtobasket.gif" alt="Add To Basket" width="221" height="30"></a></div>
					</div>
				</div>
			  </FORM>
	         
	         <a href="/shop/delivery.asp">
	         
	         <img alt="Free UK Delivery" src="../../images/shop/Free-Delivery.gif" width="135" height="122" style="float: right;" class="style1"></a>
	         <h4 style="margin-left: 5px;">
	         <%=aryRows(1, 0)%> Description</h4>
	         
	         <DIV align="justify" style="margin-left: 5px; margin-right: 5px; clear: left;">
	         <%
	         Response.Write ChangeMacros(strOutDB(aryRows(16,0)))
	         Response.Write ChangeMacros(strOutDB(aryRows(15,0)))
	         %>
	         </DIV>
			<%
			 If blnPreorder Then
			 	Call DisplayPreOrderInfo(blnDoOption)
			 End If
			%>
		<%Else%> 
			<%
			'Get rid of HTML here... 
			strBody = strOutDB(aryRows(13,i))
			strBody = Replace(strBody, "<BR>", "##BR##", 1, -1, 1)
			strBody = stripHTML(strBody)
			strBody = Replace(Server.HTMLEncode(strBody), "##BR##", "</p><p>", 1, -1, 1)
			If UBound(aryExtraNormalImages, 2) >= 0 Then
				%>
				<p><img src="http://www.drinkstuff.com/productimg/<%=strOutDB(aryExtraNormalImages(0,i))%>.<%=strOutDB(aryExtraNormalImages(1,i))%>" alt="." width="150"></p>
			<%End If%>
			<p><%=strBody%></p>
		    <%If blnDoOption Then%>
		    <p>
		    	<%For j=0 To UBound(aryRows,2)%>
		    		<%
					Call GetProductVerPrice(aryRows, j, False, dblPrice, strAffCaption)
					dblPrice = FormatNumber(dblPrice,2)
			    	%>
					<%If Trim(aryRows(3,j)) <> "" Then%><%=Server.HTMLEncode(stripHTML(strOutDB(aryRows(3,j))))%><%Else%>Default<%End If%> - <%=strAffCaption%> <b> &pound;<%=dblPrice%></b><br/>
				<%Next%>
			</p>
		    <%Else%>
		    	<%
				Call GetProductVerPrice(aryRows, 0, False, dblPrice, strAffCaption)
		    	%>
		      	<p align="left"><span class=price><%=strAffCaption%> &pound;<%=dblPrice%></span><%If Trim(aryRows(3,i)) <> "" Then%>
			   (<%=strOutDB(aryRows(3,i))%>)<%End If%></p>
				<p><%if blnCollectionOnly Then%>(Collection only)<br/><%End If%></p>
		    <%End If%> 
		    <p>To buy, please visit <a href="http://www.cocktail.uk.com/shop/products/product.asp?ID=<%=Request("ID")%>">www.cocktail.uk.com</a>.</p>
		    <%End If%><%
			End If
		End If
	End Sub
	
	Public Sub GetProductVerPrice(aryRows, j, blnAddVAT, dblPrice, strAffCaption)
		Dim dblDiscount
		
		dblPrice = aryRows(2,j)
  		If blnPricesInclVAT AND m_intVATChargeable = 0 then
  			dblPrice = aryRows(6,j)
  		End If
  		
        strAffCaption = ""
		If aryRows(7,j) <> "" And IsNumeric(aryRows(7,j)) Then
			If (aryRows(8,j) <> "" And IsDate(aryRows(8,j))) OR IsNull(aryRows(8,j)) Then
				If Now() < aryRows(8,j) OR IsNull(aryRows(8,j)) Then
					dblPrice = FormatNumber(aryRows(7,j),2)
					strAffCaption = "<font color=""red""><STRIKE>Was "&strCurrencySymbol&FormatNumber((aryRows(2,j)/dblCurrencyFactor),2)&"</STRIKE></font> now"
					If blnPricesInclVAT AND m_intVATChargeable = 0 then
						dblPrice = aryRows(9,j)
					End If
				End If
			End If 
		End If
		
		If blnAddVAT AND NOT blnPricesInclVAT Then
			dblPrice = dblPrice * (1+(aryRows(10,j)/100))
		End If
	End Sub
									
	Public Sub GetProductVerPriceCat(aryRows, j, blnAddVAT, dblPrice, strAffCaption)
		Dim dblDiscount

		dblPrice = aryRows(4,j) 
  		If blnPricesInclVAT AND m_intVATChargeable = 0 then
  			dblPrice = aryRows(8,j)
  		End If
  		
        strAffCaption = "Price"
		If aryRows(10,j) <> "" And IsNumeric(aryRows(10,j)) Then
			If (aryRows(11,j) <> "" And IsDate(aryRows(11,j))) OR IsNull(aryRows(11,j)) Then
				If Now() < aryRows(11,j) OR IsNull(aryRows(11,j)) Then
					dblPrice = FormatNumber(aryRows(10,j),2)
					strAffCaption = "<font color=""red""><STRIKE>Was "&strCurrencySymbol&FormatNumber((aryRows(3,j)/dblCurrencyFactor),2)&"</STRIKE></font> now"
					If blnPricesInclVAT AND m_intVATChargeable = 0 then
						dblPrice = aryRows(12,j)
					End If
				End If
			End If 
		End If
		
		If blnAddVAT AND NOT blnPricesInclVAT Then
			dblPrice = dblPrice * (1+(aryRows(13,j)/100))
		End If
	End Sub
									
	Private Sub DisplaySubCategories()
		Dim i
		strSQL = "DS_GETSUBCATS @catID="&m_intCategory
		response.write strSQL
		rs.open strSQL, cn
		If NOT rs.EOF Then
			m_blnGotSubCats = True
			i=0
			%><div class="row"><%
			WHILE NOT rs.EOF
				i=i+1
				%><div class=large-4 column small-6 style="padding-bottom:25px;"><A href="<%=Server.URLEncode(strOutDB(rs("url")))%>.asp" style="font-weight: 400; font-size: 130%">
				<%Call GetSubCategoryImage(rs("ID"))%><%=strOutDB(rs("name"))%>&nbsp;&raquo;</A></div>
				<%rs.movenext
			Wend
			%>
			</div>
			<%
		End If
		rs.close
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
	
	Private Sub DisplayPreOrderInfo(blnMany)
	%>
		<table border="0" cellpadding="2" style="border-collapse: collapse" width="100%">
			<tr>
				<td bgcolor="#636388" background="/images/breadcrumbbg.gif">
				<p align="left"><font color="#FFFFFF"><b>Information on pre-ordering</b></font></td>
			</tr>
			<tr>
				<td><%If blnMany Then%>Some or all of the items above are out of stock and can be 
				pre-ordered.<%Else%>This item is out of stock, but good news, it can be pre-ordered.<%End If%> This means we will take payment for the item now and you 
				will be allocated the item as soon as it comes into stock. Of course, we 
				won't charge you a delivery fee, and any other items you order that 
				are in stock will be shipped out separately so you'll get all your items 
				as soon as possible. If you decide you don't want pre-ordered items any 
				more, please <a href="http://www.awin1.com/awclick.php?gid=73406&mid=8&awinaffid=176043&linkid=101500&clickref=&p=/contactus.asp">contact us</a> ASAP so we can cancel the order.</td>
			</tr>
		</table>
		<br>
	<%
	End Sub

End Class
%>