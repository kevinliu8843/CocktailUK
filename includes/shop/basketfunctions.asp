<%
Function addbasketElement( aryBasket, numberItems )
' this function will return the *index* of the next
' available slot in the basketArray. If the basketArray 
' is full, it is expanded before the index is returned!
'
   If numberItems >= UBound(aryBasket,2) Then
        ' add 10 elements!
        ReDim Preserve aryBasket( BASKET_COLUMNS, numberItems + 10 )
    End If
    numberItems = numberItems + 1
    addbasketElement = numberItems
End Function

Function getNumberItems(aryBasket)
'Returns the number of different items in the basket
'(not total number ot items)
	Dim i,j
	j=-1
	If NOT IsArray(aryBasket) then
		Exit Function
	End If
	For i=0 To UBound(aryBasket,2)
		If Int(aryBasket(ITEM_QUANTITY,i)) > 0 Then
			j=j+1
		End If
	Next
	getNumberItems = j
End Function

Function getTotalNumberItems(aryBasket)
'Returns the total number of items in the basket
	Dim i,j
	intTotalNumberItems = 0
	If NOT IsArray(aryBasket) then
		Exit Function
	End If
	For i=0 To UBound(aryBasket,2)
		If Int(aryBasket(ITEM_QUANTITY,i)) > 0 Then
			intTotalNumberItems = intTotalNumberItems + Int(aryBasket(ITEM_QUANTITY,i))
		End If
	Next
	getTotalNumberItems= intTotalNumberItems 
End Function

Private Sub RemoveItem(item, aryBasket, numberItems)
'Removes a single item from the basket
	Dim numberItemsDup, aryBasketDup, prodNum, intitem, j, iloop
	
	If NOT IsArray(aryBasket) then
		Exit Sub
	End If

	If IsNumeric(item) then
		intitem = Int(item)
	End If

	numberItemsDup = numberItems
	aryBasketDup = aryBasket
	numberItems = -1
	prodNum = -1
	For iloop=0 To numberItemsDup
		If intitem <> iloop Then
			prodNum = prodNum + 1
			For j=0 to UBound(aryBasket,1)
				aryBasket(j,prodNum) = aryBasketDup(j,iloop)
			Next
			numberItems = numberItems + 1
		End If
	Next
End Sub

Private Sub tidyBasket(aryBasket)
'Removes items from the basket that have a quantity >= 0
	Dim numItems, iremove
	If NOT IsArray(aryBasket) then
		Exit Sub
	End If
	numItems = UBound(aryBasket,2)
	iremove = 0
	For i=0 To numItems
		If Int(aryBasket(ITEM_QUANTITY,iremove)) <= 0 OR (Session("platinummember") AND aryBasket(ITEM_PRODOFFERID,iremove) > 0) Then
			call RemoveItem(iremove, aryBasket, numItems)
			iremove = iremove - 1
		End If
		iremove = iremove + 1
	Next
End Sub

Sub emptyBasket(aryBasket)
'Removes all items from the basket and resets some session vars.
	Dim basketitems
	If NOT IsArray(aryBasket) then
		Exit Sub
	End If
	For basketitems=0 To UBound(aryBasket,2)
		aryBasket(ITEM_PRODVERID,basketitems) = ""
		aryBasket(ITEM_QUANTITY,basketitems) = 0
	Next
	Session("pointsadded") = False
	Session("pointstouse") = 0
	Session("intOrderID") = ""
	Session("numberItems") = 0
	Session("totalNumberItems") = 0
	Session("valueItems") = 0
End Sub

Sub setupProducts(aryBasket, numberItems, cn, rs, aryProducts, blnStockCheck, blnStockFlag)
'Sets up the aryProducts array
'Does stock checking (if required)
'Sets up the unit prices, total prices, quantity, vat rate etc... for each product in the basket

	Dim rsDel, blnAllowed, blnRemoved, k, strName, intStock, dteDueIn, strDateDueIn, blnPreOrder, intMaxPreorder, intStockStatus
	blnStockFlag	= FALSE
	Set rsDel= Server.CreateObject("ADODB.Recordset")
	Set objProd = New CProduct 'Needed to check for affiliate special prices
	k=-1
	For i=0 To numberItems
		strSQL = "SELECT  dsproduct.name, dsproductver.ID AS prodverid, dsproduct.ID AS prodid, dsproductver.price, dsproductver.subtext, dsproductver.vat, dsimage.ID as imgid, dsimage.type, saleprice, saleexpires"
		strSQL = strSQL & " FROM         dsproductver INNER JOIN"
		strSQL = strSQL & "                      dsproduct ON dsproductver.prodID = dsproduct.ID"
		strSQL = strSQL & "           INNER JOIN dsimage ON dsproduct.ID = dsimage.prodID "
		strSQL = strSQL & " WHERE     (dsproductver.status = 1) AND (dsproduct.status = 1) AND (dsproductver.ID = "&aryBasket(0,i)&") AND (dsimage.imagesize=0)"
		rs.open strSQL, cn, 0, 3
		If NOT rs.EOF Then
			blnRemoved = FALSE
			numProducts = numProducts + 1
			k = k + 1
			ReDim Preserve aryProducts(BASKET_COLUMNS+1, numProducts)
			
			strName = strOutDB(rs("name"))
			If strOutDB(rs("subtext")) <> "" then
				strName = strName & " (" & strOutDB(rs("subtext")) & ")"
			End If

			If blnStockCheck Then
				'Do stock checking here...
				rsDel.open "EXECUTE DS_GETRAWPRODDETAILS @prodverID="&aryBasket(0,i), cn, 0, 3
				intStock = rsDel("stock")
				dteDueIn = rsDel("dateduein")
				intStockStatus = Int(rsDel("stockstatus"))
				If rsDel("stock") <= 0 AND rsDel("preorder") Then
					blnPreOrder = True
					intMaxPreorder = rsDel("maxpreorder")
				Else
					blnPreOrder = False
				End If
				rsDel.close
				If NOT IsNumeric(intStock) Then
					intStock = 0
				End If

				aryBasket(ITEM_PREORDER,i) = 0
				If intStockStatus = PRODUCT_STOCK AND Int(intStock) < Int(aryBasket(ITEM_QUANTITY,i)) AND Int(intStock) > 0 Then
					addnote("Sorry, you required "&aryBasket(ITEM_QUANTITY,i)&" items of """&strName&""", but we only have "&intStock&" available.<BR/>If you require more please call or <A href=""/services/contact.asp"">email</A> as more stocks may be arriving any day")
					aryBasket(ITEM_QUANTITY,i) = intStock
					blnStockFlag = True
					blnRedirectBack = FALSE
				End If
				If intStockStatus = PRODUCT_STOCK Then
					If Int(intStock) <= 0 AND NOT blnPreOrder Then
						If dteDueIn <> "" Then
							strDateDueIn = " We are expecting more in on the "&StripThisYear(MediumDate(dteDueIn))&"."
						Else
							strDateDueIn = ""
						End If
						addnote("Sorry, we have no stock of the """&strName&""" at the moment."&strDateDueIn&"<br/>&nbsp;<br/>If you like, we can notify you when it is back in stock. <A HREF=""http://www.drinkstuff.com/products/affiliate.asp?ffID=10724&page=/products/stock_notification.asp?ID="&rs("prodID")&""">Click here</A> to find out more.")
						aryBasket(ITEM_QUANTITY,i) = 0
						blnRemoved = True
						blnStockFlag = True
						blnRedirectBack = FALSE
					Else
						If Int(intStock) <= 0 AND blnPreOrder Then
							aryBasket(ITEM_PREORDER,i) = 1
							If Int(aryBasket(ITEM_QUANTITY,i)) > intMaxPreorder Then
								addnote("Sorry, you can only pre-order a maximum of "&intMaxPreorder&" """&strName&""" at the moment.")
								aryBasket(ITEM_QUANTITY,i) = intMaxPreorder
								blnRedirectBack = FALSE
							End If
						End If
					End If
				End If
			End If
			
			If blnApplyDeliveryZoneRestrictions Then
				'Get allowed delivery zones here...
				strSQL = "SELECT dsdelivery.ID FROM dsprodallowdelivery INNER JOIN dsproduct ON dsprodallowdelivery.prodid = dsproduct.ID INNER JOIN dsdelivery ON dsprodallowdelivery.delid = dsdelivery.ID WHERE dsdelivery.status=1 AND dsprodallowdelivery.prodid=" & rs("prodid")
				rsDel.open strSQL, cn, 0, 3
				blnAllowed = False
				If NOT rsDel.EOF Then
					While NOT rsDel.EOF 
						If CStr(session("delivery")) = CStr(rsDel("ID")) Then
							blnAllowed = True
						End If
						rsDel.movenext
					Wend
				Else
					blnAllowed = True
				End If
				rsDel.close
				If NOT blnAllowed Then 
					If NOT blnRemoved Then
						addnote("Sorry, the "&strName&" is not available for delivery in your selected delivery region. We have removed it from your basket for you.")
						aryBasket(ITEM_QUANTITY,i) = 0
						numProducts = numProducts - 1
						k = k - 1
					End If
					blnRemoved = True
				End If
				
				'Is it collection only?
				If IsCollectionOnly(cn, rsDel, rs("prodid")) Then
					If NOT Session("admin") Then	
						If NOT blnRemoved Then
							addnote("Sorry, the "&strName&" is only available for collection or local delivery. Call 01223 872769 to check if we can deliver to your area. We have removed it from your basket for you.")
							aryBasket(ITEM_QUANTITY,k) = 0
							numProducts = numProducts - 1
							k = k - 1
						End If
						blnRemoved = True
					Else 
						addnote("Admin note: """&strName&""" is normally only available for collection")
					End If
				End If
			End If

			If NOT blnRemoved Then
				aryProducts(ITEM_NAME,k) = strName
				aryProducts(ITEM_QUANTITY,k) = aryBasket(ITEM_QUANTITY,i)
				aryProducts(ITEM_PRICE,k) = FormatNumber(CalculateGrossFromNet(rs("price"), rs("vat")), 2)
				If rs("saleprice") <> "" And IsNumeric(rs("saleprice")) Then
					If (rs("saleexpires") <> "" And IsDate(rs("saleexpires"))) OR IsNull(rs("saleexpires")) Then
						If Now() < rs("saleexpires") OR IsNull(rs("saleexpires")) Then
							aryProducts(ITEM_PRICE,k) = CalculateGrossFromNet(rs("saleprice"), rs("vat"))
							aryProducts(ITEM_NAME,k) = aryProducts(ITEM_NAME,k) & " - on sale, was " & FormatNumber(rs("price"),2)
						End IF
					End If 
				End If
				aryProducts(TOTAL_PRICE,k) = FormatNumber(CSng(aryProducts(ITEM_QUANTITY,k)) * CSng(aryProducts(ITEM_PRICE,k) ), 2)
				aryProducts(ITEM_LINK,k) = rs("prodid")
				aryProducts(ITEM_PRODVERID,k) = aryBasket(0,i)
				aryProducts(VAT_RATE,k) = rs("vat")
				aryProducts(ITEM_IMAGE,k) = CStr(rs("imgID")) & "." & strOutDB(rs("type"))
				If objProd.IsAffOffer(aryBasket(0,k), dblPrice) Then
					aryProducts(ITEM_PRICE,k) = FormatNumber(CalculateGrossFromNet(dblPrice, rs("vat")), 2) 'FormatNumber(dblPrice,2)
					aryProducts(TOTAL_PRICE,k) = FormatNumber(CSng(aryProducts(ITEM_QUANTITY,k)) * CSng(dblPrice), 2) 'FormatNumber(dblPrice,2)
					aryProducts(ITEM_NAME,k) = aryProducts(ITEM_NAME,k)' & " (" & Session("affSiteName") & " price)"
				End If
			End if
		Else
			aryBasket(ITEM_QUANTITY,i) = 0
		End If
		rs.close
	Next
	Set objProd = Nothing
	Set rsDel = Nothing

	'Do a tidy up to remove 0 or less quantity items
	call tidyBasket(aryBasket)
	
	'Add free product here if necessary
	'DumpOut()
	numberItems = getNumberItems(aryBasket)
	'DumpOut()
End Sub

Sub addPromoCode(rs, dblDiscountPerc, dblDiscount, strOut, aryProducts, numberItems)
'Adds a promotional code if required
	If Session("promo") <> "" Then
		strSQL = "SELECT * from dspromo WHERE code='"&strIntoDB(Session("promo"))&"' AND status=1"
		rs.open strSQL, cn, 0, 3
		If NOT rs.EOF Then
			If GetSubTotal(aryProducts, numberItems) >= CDbl(rs("minspend")) Then
				dblDiscountPerc = CSng(rs("percentage"))
				dblDiscount     = CSng(rs("code_value"))
				Session("intOrderID") = ""
			Else
				strOut 			= "Sorry, you need to spend at least "&FormatNumber(rs("minspend"),2)&" to use this voucher"
				Session("promo")= ""
				dblDiscountPerc = 0
				dblDiscount     = 0
				Session("intOrderID") = ""
			End If
		Else
			strOut 			= "Sorry, the promotional code you entered is invalid."
			Session("promo")= ""
			dblDiscountPerc = 0
			dblDiscount     = 0
			Session("intOrderID") = ""
		End If
		rs.close
	End If
End Sub

Function addnote(strString)
	strNotes = strNotes & strString & "<P>"
End Function

Sub calculateCosts(rs, cn, totalValue, numberItems, deliveryCost, deliveryFree, deliveryMinValue, blnNextStage, discountValue, dblVAT, intPointsToUse)
'Determines the costs, taking into account any reward points and/or promotional codes
	Dim dblVATPercentage

	totalValue = GetSubTotal(aryProducts, numberItems)
	If totalValue > 0 AND numberItems > -1 then
		If deliveryFree > 0 Then
			If totalValue >= deliveryFree Then
				deliveryCost = CSng(0)
				strDeliveryNote = "Well done! You have qualified for free delivery"
			Else
				strDeliveryNote = "<B>Tip: </B>If you spend a further "&FormatNumber(deliveryFree-totalValue,2)&", you will qualify for free delivery"
			End If
		End If
		If deliveryMinValue > 0 AND totalValue < deliveryMinValue Then
			addnote("Unfortunately, we only deliver to this region when you spend at least "&FormatNumber(deliveryMinValue,2)&" (excl. delivery) with us.")
			If InStr(LCase(Request.ServerVariables("SCRIPT_NAME")), "stage2.asp") > 0 then
				addNote("<SCRIPT Language=Javascript>location.href='../basket.asp'</SCRIPT>")
			End If
			blnNextStage = FALSE
		End If
		
		'Have products less that 1 free delivery (takes care of 99p chip set + free delivery)
		'response.write totalOld
		If totalValue < 1.00 AND totalValue > 0 Then
			deliveryCost = 0
		End if
		
		'Apply promotional code discounts here...
		If Session("promo") <> "" then
			totalOld = totalValue 
			totalValue = totalValue - dblDiscount
			totalValue = totalValue *(100-dblDiscountPerc)/100
			totalValue = Max(totalValue,0) ' Need a minimum price of 0 pounds (i.e. no refunds possible from promo codes! EEK)
			discountValue= FormatNumber(-1*(totalValue - totalOld ),2)
		Else
			discountValue= 0
		End If
		
		'Do reward points thing...
		intPoints = 0
		If Session("memID") <> "" Then
			If IsNumeric(Session("memID")) Then
				strSQL = "SELECT points from dsmember WHERE ID=" & Session("memID")
				rs.open strSQL, cn ,0, 3
				If NOT rs.EOF Then	
					intPoints = Int(rs("points"))
				Else
					intPoints = 0
				End If
				rs.close
			End If
		End If

		intPointsToUse = Max(Int(Min(intPoints, totalValue)),0)

		'Apply reward points here...
		If Session("pointsadded") Then
			If intPointsToUse > 0 Then
				Session("pointstouse") = intPointsToUse
				discountValue = discountValue + intPointsToUse
				discountValue = FormatNumber(discountValue, 2)
				totalValue = totalValue - intPointsToUse
			Else
				Session("pointsadded") = False
				Session("pointstouse") = 0
			End If
			Session("intOrderID") = ""
		Else
			Session("pointstouse") = 0
			Session("intOrderID") = ""
		End If

		totalValue = totalValue + deliveryCost
	End If
	
	totalValue = FormatNumber(totalValue,2)
	deliveryCost= FormatNumber(deliveryCost,2)
	If totalValue <= 0 Then
		deliveryCost = 0
	End If
End sub

Function GetSubTotal(aryProducts, numberItems)
'Determines the sub total of the products in the basket
	Dim totalValue
	For i=0 To numberItems
		totalValue = totalValue + CSng(aryProducts(TOTAL_PRICE,i))
	Next
	GetSubTotal = totalValue
End Function

Function DumpOut()
	Dim i
	If Session("admin") Then
		Response.write "<P>Products Array<BR>"
		For i=0 To Ubound(aryProducts,2)
			response.write "Name: " & aryProducts(ITEM_NAME,i) & ", Quantity: " & aryProducts(ITEM_QUANTITY,i) & "<BR>"
		Next
	End If
End Function

Function GetAlsoBought(intProdID, strProductName)
'GetAlsoBought = ""
'Exit Function
	Dim aryExtraProds(3,4), intExtraProds, i, strOut
	call GetAlsoBoughtData(intProdID, aryExtraProds, intExtraProds)
	If intExtraProds > 0 Then
              strOut = strOut & "<TABLE border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""4"" background=""/images/grad_write_purple.gif"">"
                strOut = strOut & "<TR>"
                  strOut = strOut & "<TD bgcolor=""#747495"" height=""1"" background=""/images/breadcrumbbg.gif""><FONT color=""#FFFFFF""><B>Customers who bought <I>"&FormatProductName(strProductName)&"</I> also bought...</B></FONT></TD>"
                  strOut = strOut & "</TR>"
                  strOut = strOut & "<TR>"
                  strOut = strOut & "<TD height=""2"">"
                  strOut = strOut & "<TABLE border=""0"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"" width=""100%"" id=""AutoNumber4"">"
                    strOut = strOut & "<TR>"
                      For i=0 To intExtraProds-1
                        strOut = strOut & "<TD valign=""top"" width=" & CStr(100/intExtraProds) & "% align=""center""><A href=""/shop/products/product.asp?ID="&aryExtraProds(0,i)&"""><IMG src=""/shop/getimage.asp?img="&aryExtraProds(3,i)&""" border=""0"" height=""25""><BR><font size=""1"">"&aryExtraProds(1,i)&"</font></A><br><font size=""1"">"&aryExtraProds(2,i)&"</font></TD>"
                      Next
                    strOut = strOut & "</TR>"
                  strOut = strOut & "</TABLE>"
                  strOut = strOut & "</TD>"
                strOut = strOut & "</TR>"
                strOut = strOut & "<TR>"
                  strOut = strOut & "<TD bgcolor=""#747495"" height=""2""></TD>"
                strOut = strOut & "</TR>"
                strOut = strOut & "</TABLE>"
              strOut = strOut & "<IMG border=""0"" src=""../img/pixel.gif"" width=""5"" height=""5"">"
	End If
	GetAlsoBought = strOut
End Function

Sub GetAlsoBoughtData(intProdID, aryExtraProds, intExtraProds)
	on error resume next
	Dim objXmlHttpCat, objData, strXML, objXmlDoc, objExtraProds, i
	If intProdID = "" Then
		Exit Sub
	End If
	Set objXmlHttpCat= Server.CreateObject("MSXML2.ServerXMLHTTP")
	objXmlHttpCat.open "GET", "http://www.drinkstuff.com/affiliate/alsobought.asp?prodID="&intProdID , False
	objXmlHttpCat.send ""
	Set objXmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	strXML = objXmlHttpCat.responseXml.xml
	objXmlDoc.loadXML(strXML)
	Set objXmlHttpCat = nothing
	Set objExtraProds= objXmlDoc.getElementsByTagName("ALSOBOUGHT/NUMPRODUCTS")
	intExtraProds = Int(objExtraProds.item(0).text)
	For i=0 To intExtraProds-1
		Set objData = objXmlDoc.getElementsByTagName("ALSOBOUGHT/PRODUCT"&i+1&"/PRODID")
		aryExtraProds(0,i) = objData.item(0).text
		Set objData = objXmlDoc.getElementsByTagName("ALSOBOUGHT/PRODUCT"&i+1&"/PRODNAME")
		aryExtraProds(1,i) = objData.item(0).text
		Set objData = objXmlDoc.getElementsByTagName("ALSOBOUGHT/PRODUCT"&i+1&"/PRODPRICE")
		aryExtraProds(2,i) = objData.item(0).text
		Set objData = objXmlDoc.getElementsByTagName("ALSOBOUGHT/PRODUCT"&i+1&"/PRODIMAGE")
		aryExtraProds(3,i) = objData.item(0).text
	Next

	Set objData = nothing
	Set objXmlDoc = nothing
End Sub

Function FormatProductName(strName)
	If Len(strName) > 23 Then
		FormatProductName = Left(strName, 23) & "..."
	Else
		FormatProductName = strName
	End If
End Function
%>