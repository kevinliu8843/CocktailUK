<% Option Explicit %>
<%
strTitle 	= "Your basket"
dim cn
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/basketfunctions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<%
Session.Timeout = 15

Dim intDelivery, strNotes, aryBasket, i, X, aryProducts(), intQuantity
Dim blnBasketEmpty, strDeliveryZones , numberItems, intPVID
Dim basketMaxUsed, basketItem, numProducts, totalValue, deliveryCost
dim deliveryFree, strDeliveryNote, deliveryMinValue, blnAddItem
Dim blnRedirectBack, strAdd, strTo, strRedirect, dblDiscountPerc, dblDiscount
Dim discountValue, totalOld, blnAffMode, dblPrice, objProd, blnNextStage
Dim strFreeGoodsnote, strQuantites, strProducts, blnStockCheck, strOut
Dim strContinueShopping, blnStockFlag, intTotalNumberItems, dblBasketSubTotal
Dim intPoints, intPointsToUse, blnPointsAdded, intPointsUsed
Dim dblVAT, intLastProduct, strAlsoBought, blnItemAddedToBasket

ReDim aryProducts(BASKET_COLUMNS+1, 0)
numProducts = -1

'Affiliate prices setup here...
blnAffMode = (Session("affid") <> "")

'Delivery zone set up
If Session("delivery") = "" Then
	Session("delivery") = 1
End If
If Request("delivery") <> "" AND IsNumeric(Request("delivery")) Then
	Session("delivery") = Int(Request("delivery"))
End If
intDelivery = Int(Session("delivery"))

'Set up Session variable "basket"
aryBasket = Session("basket")
numberItems  = Session("numberItems")
numberItems = getNumberItems(aryBasket) 'More reliable?

blnItemAddedToBasket = False
blnRedirectBack = FALSE

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open strDB

'======================='Add items here'=======================
'======================='Add items here'=======================
'======================='Add items here'=======================
If IsNumeric(Request("prodverID")) AND IsNumeric(Request("quantity")) AND Request("prodverID") <> "" AND Request("quantity") <> "" Then
	If Request("quantity") > 0 Then
		intQuantity	= Int(Request("quantity"))
		intPVID		= Int(Request("prodverID"))
		Call AddProd(intPVID, intQuantity)
	End If
Else
	For Each X in Request.Form
		If Left(X, 8) = "quantity" Then
			intPVID = Replace(LCase(X), "quantity", "")
			If IsNumeric(intPVID) AND IsNumeric(Request(X)) Then
				intQuantity	= Int(Request(X))
				Call AddProd(intPVID, intQuantity)
			End If
		End If
	Next
End If

'=======================Recalculate values here...===========================
'=======================Recalculate values here...===========================
'=======================Recalculate values here...===========================
If Request("recalculate") ="true" Then
	For i=0 To numberItems
		If IsNumeric(Request("quantity"&i)) AND UBound(aryBasket, 2) >= i Then
			aryBasket(ITEM_QUANTITY, i) = Int(Min(CDbl(Request("quantity"&i)),10000))
		End If
	Next
End If

'======================='Remove individual products here'=======================
'======================='Remove individual products here'=======================
'======================='Remove individual products here'=======================
If Request("remove") <> "" AND IsNumeric(Request("remove")) Then
	If UBound(aryBasket, 2) >= Int(Request("remove")) Then
		aryBasket(ITEM_QUANTITY,Int(Request("remove"))) = 0
	End If
End If

'======================='Clear basket here'=======================
'======================='Clear basket here'=======================
'======================='Clear basket here'=======================
If Request("emptybasket") = "true" Then
	call emptyBasket(aryBasket)
	addnote("Your basket has been emptied")
End If

'Do a tidy up to remove 0 or less quantity items
call tidyBasket(aryBasket)
numberItems = getNumberItems(aryBasket)

'======================='Set up aryProducts here...'=======================
'======================='Set up aryProducts here...'=======================
'======================='Set up aryProducts here...'=======================
blnStockCheck = True
call setupProducts(aryBasket, numberItems, cn, rs, aryProducts, blnStockCheck, blnStockFlag)

blnBasketEmpty = ( numberItems < 0 )

'========================='(Re) Set up session variable and local page vars'=========================
'========================='(Re) Set up session variable and local page vars'=========================
'========================='(Re) Set up session variable and local page vars'=========================
Session("basket") = aryBasket
Session("numberItems") = numberItems
Session("totalNumberItems") = getTotalNumberItems(aryBasket)

'========================='Set up promotional code here...'=========================
'========================='Set up promotional code here...'=========================
'========================='Set up promotional code here...'=========================
If Request("recalculate") = "true" Then
	Session("promo") = Request("promo")
End If
call addPromoCode(rs, dblDiscountPerc, dblDiscount, strOut, aryProducts, numberItems)
If strOut <> "" Then
	call addnote(strOut)
End If

blnNextStage = TRUE

call calculateCosts(rs, cn, totalValue, numberItems, deliveryCost, deliveryFree, deliveryMinValue, blnNextStage, discountValue, dblVAT, intPointsToUse)

Session("valueItems") = FormatNumber(totalValue-deliveryCost,2) 'Remove delivery cost when displaying basket in RH corner

'========================='Set up extra also bought products here...'=========================
'========================='Set up extra also bought products here...'=========================
'========================='Set up extra also bought products here...'=========================
If NOT blnBasketEmpty Then
	intLastProduct = UBound(aryProducts, 2)
	strAlsoBought = GetAlsoBought(aryProducts(ITEM_LINK, intLastProduct), aryProducts(ITEM_NAME, intLastProduct))
End If

cn.close
Set cn = nothing
Set rs = nothing

'========================='Redirect back to the previous refering page here...'=========================
'========================='So user does stay on basket page when they add a product.'===================
'=======================================================================================================
blnRedirectBack= FALSE
If blnRedirectBack Then
	If Request("previous") <> "" Then
		strTo = strTo & "&previous="
		If InStr(Request("previous"), "previous=") > 0 Then
			strTo = strTo & Left(Request("previous"), InStr(Request("previous"), "&previous=")-2)
		Else
			strTo = strTo & Request("previous")
		End If
	End If
	'Strip previous and added request fields (In case they click "add" twice on same page)
	strRedirect = Replace(Replace(Request.ServerVariables("HTTP_REFERER"),"&added=true",""),"?added=true","")
	If InStr(strRedirect, "&previous=") > 0 Then
		strRedirect = Left(strRedirect, InStr(strRedirect, "&previous=")-1)
	End If
	If InStr(strRedirect, "?") <= 0 Then
		strAdd = "?"
	Else
		strAdd = "&"
	End If
	strRedirect = strRedirect & strAdd & "added=true" & strTo
	'response.write strRedirect
	Response.redirect(strRedirect)
End If

strContinueShopping = Request("referer")
%>
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" --><%If NOT blnBasketEmpty Then%>
<script language="JavaScript">
function nextstage()
{
	<%If NOT blnNextStage Then%>
		alert("Unfortunatley, we only deliver to this region when you spend at least <%=FormatNumber(deliveryMinValue,2)%> with us.");
		return false;
	<%else%>
		//alert("You will now be transferred to Drinkstuff.com, our partner site, ready for payment.\nYou can still add/remove products when there.");
		return true;
	<%End if%>
}
function showApplyImg()
{
	strOriginal = "<%=Session("promo")%>"
	if (document.basketform.promo.value != strOriginal)
	{
		document.all.applyimg.src = "/images/shop/basket/apply_blink.gif";
	}
	else
	{
		document.all.applyimg.src = "/images/shop/basket/apply.gif";
	}
}
function showUpdateImg()
{
	document.all.update.src = "/images/shop/basket/recalculate_flash.gif";
}
function basketpopup()
{
   window.open("http://www.drinkstuff.com/member/basketpopup.asp?img=http://www.cocktail.uk.com/images/topLogoOriginal.gif&sitename=cocktail : UK","basketpopup","height=220, width=350, menubar=0, status=0")
}
<%If Session("basketpopup") <> "false" Then%>
   <%Session("basketpopup") = "false"%>
//basketpopup()
<%End If%>
</script>
<h2>Your basket</h2>
<%If blnItemAddedToBasket Then%>
<p align="center"><i><font color="#FF0000"><b>Product<%If request("quantity") <> "1" Then%>s<%End If%> 
added to your basket</b></font></i></p>
<%End If%>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td valign="top">
    <p align="justify">Since opening the shop in 1999, we have taken over 
	250,000 orders online! That's a lot of satisfied customers.</td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
  <tr>
    <td>
    <p align="center"><b><font color="red"><%=strNotes%></font></b></p>
    </td>
  </tr>
  <tr>
    <td>
    <table border="0" cellpadding="0" cellspacing="0" width="100%">
      <form method="POST" action="basket.asp" style="padding: 0" name="basketform">
       <tr>
         <td colspan="2">
         <input type="hidden" name="recalculate" value="true">
         <table border="0" width="100%" cellspacing="0" cellpadding="4" id="table3" style="background-image: url('../images/grad_write_purple_small.gif'); background-position: bottom left; background-repeat: repeat-x;">
           <tr>
             <td bgcolor="#747495" colspan="2" background="../images/breadcrumbbg.gif"><font color="#FFFFFF"><b>Item</b></font></td>
             <td align="center" bgcolor="#747495" width="15%" background="../images/breadcrumbbg.gif">
             <font color="#FFFFFF"><b>Quantity</b></font></td>
             <td align="right" bgcolor="#747495" nowrap width="15%" background="../images/breadcrumbbg.gif">
             <font color="#FFFFFF"><b>Total Price</b></font></td>
           </tr>
           <%For i=0 To numberItems%>
           <tr>
             <td bordercolor="#E7E7E7" width="3%" align="center" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
             <a style="text-decoration: none" title="View '<%=aryProducts(ITEM_NAME,i)%>'" href="/shop/products/product.asp?ID=<%=aryProducts(ITEM_LINK,i)%>">
             <img src="/shop/getimage.asp?img=<%=aryProducts(ITEM_IMAGE,i)%>" height="25" border="0" /></a></td>
             <td bordercolor="#E7E7E7" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
             <a style="text-decoration: none" title="View '<%=aryProducts(ITEM_NAME,i)%>'" href="/shop/products/product.asp?ID=<%=aryProducts(ITEM_LINK,i)%>">
             <%=aryProducts(ITEM_NAME,i)%></a> <br>
             <%If aryBasket(ITEM_PREORDER,i) = 1 Then%>
	                  <font color="#999999">
	                  <span class="smalltext"><nobr><font size="1">(<img border="0" src="../images/shop/gotstock.gif"> Pre-order - will ship separately)</font></nobr></span></font>
				  <%Else%><font size="1" color="#999999">(<img border="0" src="../images/shop/gotstock.gif"> 
             in stock)</font>
				  
				  <%End If%>
                  </td>
             <td align="center" nowrap bordercolor="#E7E7E7" width="15%" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px">
             <p class="smalltext">
             <input type="text" size="2" name="quantity<%=i%>" value="<%=aryProducts(ITEM_QUANTITY,i)%>" style="font-size: 11px">
             <a href="basket.asp?remove=<%=i%>"><font color="#FF0000" size="1">remove</font></a></p>
             </td>
             <td align="right" nowrap bordercolor="#E7E7E7" width="15%" style="border-left-width: 1px; border-right-width: 1px; border-top-width: 1px; border-bottom-style: solid; border-bottom-width: 1px"> <%=aryProducts(TOTAL_PRICE,i)%></td>
           </tr>
           <%next%>
           <tr>
             <td colspan="4" bgcolor="#747495" height="1"></td>
           </tr>
           <tr>
             <td colspan="4" height="2">
             <p align="right">
             <input src="../images/buttons/update_quantities.gif" type="image" id="update" value="Submit" border="0" alt="Click here to update your basket totals" align="middle" name="I2" width="103" height="15">
             </p>
             </td>
           </tr>
           <tr>
             <td colspan="4" bgcolor="#747495" height="1"></td>
           </tr>
           <tr>
             <td colspan="4" height="2">
             <table border="0" cellpadding="1" cellspacing="0" width="100%" height="0" id="table4">
               <tr>
                 <td align="right" rowspan="<%If Session("pointsadded") OR Session("promo") <> "" Then%>3<%Else%>2<%End If%>" valign="top">
                 <p align="left"><span class="smalltext">
                 <a style="text-decoration: none" href="javascript:displayDeliveryTimes()">
                 <font color="#FF0000">
                 <img border="0" src="../images/shop/basket/basket_arrow.gif" align="middle">View the postage 
                 &amp; packaging costs...</font></a></span></p>
                 </td>
                 <td align="right" rowspan="<%If Session("pointsadded") OR Session("promo") <> "" Then%>3<%Else%>2<%End If%>" valign="top">&nbsp;</td>
                 <td align="right" height="10" nowrap><b>
                 <font color="#FF0000" size="4"> <%=totalValue%></font></b></td>
               </tr>
             </table>
             </td>
           </tr>
           <tr>
             <td colspan="4" bgcolor="#747495" height="2"></td>
           </tr>
         </table>
         </td>
       </tr>
      </form>
      <tr valign="middle">
        <td align="center" valign="top">
        <p align="left">&nbsp;<img src="/images/shop/credit-cards.gif" align="center" /></p>
        </td>
        <td align="right" valign="top">
        <form name="continue" action="/shop/sendtodrinkstuff.asp" method="POST" onsubmit="return nextstage()">
         <p>
         <input src="../images/buttons/payment.gif" alt="Press to transfer your basket contents to Drinkstuff.com for payment" type="image" border="0" align="middle" name="I1" width="121" height="19">
         <input type="hidden" name="valueItems" value="<%=Session("valueItems")%>">
         <input type="hidden" name="from" value="basket">
         <input type="hidden" name="searchengine" value="<%=Session("searchengine")%>">
         <input type="hidden" name="keywords" value="<%=Session("keywords")%>">
         <input type="hidden" name="searchenginereferer" value="1"></p>
        </form>
        </td>
      </tr>
      <tr>
        <td colspan="2">
		<p align="center">In association with<br>
		<a href="/shop/sendtodrinkstuff.asp">
		<img border="0" src="../images/shop/drinkstuff%20logo.gif" width="200" height="26"></a>
		<img src="/images/pixel.gif" height="10"><br>
        <%=strAlsoBought%></td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<%Else%>
<h2>Your basket is empty </h2>
<table border="0" cellpadding="5" style="border-collapse: collapse" width="100%" id="table5">
  <tr>
    <td>
    <p align="center"><b><font color="red"><%=strNotes%></font></b></p>
    <p>
    <img border="0" src="/images/shop/basket/an_empty_basket.jpg" align="left"></p>
    <p>If you havent put anything in your basket then please <a href="/shop">click 
    here and we will take you to the shop</a> so you can find the items you want.</p>
    <p>If you have tried to add something to your basket but it is still empty, 
    you may not have &#39;cookies&#39; enabled on you computer. &#39;Cookies&#39; are the things 
    we use to make shopping secure when you shop at Cocktail : UK.</p>
    <p>If you previously had some items in a basket, please note that we have an 
    inactivity timeout of <%=Session.timeout%> minutes. Your basket is automatically 
    cleared then.</p>
    <p><b>Thanks,</b><br>
    The Cocktail : UK team.</p>
    </td>
  </tr>
</table>
<%End If%><!--#include virtual="/includes/shop/footer.asp" --><!--#include virtual="/includes/footer.asp" -->
<%
Sub AddProd(intPVID, intQuantity)
	If intQuantity > 0 AND intPVID > 0 Then
		'Check for existing same products and add quantities together
		blnAddItem = TRUE
		blnRedirectBack = TRUE
		For i=0 To numberItems
			If aryBasket(ITEM_NAME, i) = intPVID then
				aryBasket(ITEM_QUANTITY, i) = aryBasket(ITEM_QUANTITY, i) + intQuantity
				blnAddItem = FALSE
				blnItemAddedToBasket = True
			End If
		Next
		
		If blnAddItem Then
			Session("intOrderID") = ""
			'add prodverID and quantity to basket (check validity in a mo)
			basketItem = addBasketElement( aryBasket, numberItems )
			aryBasket(ITEM_NAME, basketItem) = intPVID 
			aryBasket(ITEM_QUANTITY, basketItem) = intQuantity
			blnItemAddedToBasket = True
		End If
	End If
End Sub
%>