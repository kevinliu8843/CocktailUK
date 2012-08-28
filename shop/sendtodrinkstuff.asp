<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/basketfunctions.asp" -->
<%
aryBasket = Session("basket")
numberItems  = Session("numberItems")
numberItems = getNumberItems(aryBasket)
totalNumberItems = Session("totalNumberItems")
delivery = Session("delivery")
strSearchEngineReferer = Request("searchenginereferer")

'Set up the products and quantities for next stage
If numberItems >= 0 Then
	For i=0 To numberItems-1
		strProducts = strProducts & aryBasket(ITEM_NAME ,i) & ","
		strQuantites = strQuantites& aryBasket(ITEM_QUANTITY,i) & ","
	Next
	strProducts = strProducts & aryBasket(ITEM_NAME ,numberItems)
	strQuantites= strQuantites& aryBasket(ITEM_QUANTITY,numberItems)
End If

strQueryString = "?products=" & strProducts & "&quantities=" & strQuantites & "&numberItems=" & numberItems & "&valueItems=" & Session("valueItems") & "&delivery=" & delivery & "&totalNumberItems=" & totalNumberItems & "&promo=" & Session("promo") & "&affID=10724&pointstouse=" & Session("pointstouse") & "&pointsadded=" & Session("pointsadded") & "&from=cocktailuk&goto=member/basket.asp&searchengine="&Session("searchengine")&"&keywords="&Session("keywords")&"&searchenginereferer="&strSearchEngineReferer 

strTitle="Transferring to drinkstuff.com"
%>
<!--#include virtual="/includes/header.asp" -->
<META http-equiv="refresh" content="2;url=http://www.drinkstuff.com/member/getbasketfromremote.asp<%=strQueryString%>" />
<SCRIPT>
function redirecttods(){
location.href = "http://www.drinkstuff.com/member/getbasketfromremote.asp<%=strQueryString%>"
}
timerid = setTimeout("redirecttods()", 2500);
</SCRIPT>
<CENTER>
<table border="0" cellspacing="0" cellpadding="2" style="margin-top: 20px;">
	<tr>
		<td valign=top nowrap>
		<img border="0" src="/images/shop/wait.gif" width="48" height="48"></td>
		<td valign=top nowrap><b>Transferring to our equipment partners drinkstuff.com</b>.<br>For payment and to confirm your order.</td>
	</tr>
<tr><td></td><td><b><A href="/shop/basket.asp">&lt; &lt; Return to Cocktail : UK</a>  | <A href="http://www.drinkstuff.com/member/getbasketfromremote.asp<%=strQueryString%>">Continue to Drinkstuff &gt; &gt;</a></b></td></tr>
</table>
</CENTER>
<P>&nbsp;</p>
<P>&nbsp;</p>
<P>&nbsp;</p>
<P>&nbsp;</p>
<P>&nbsp;</p>
<P>&nbsp;</p>
<P>&nbsp;</p>
<P>&nbsp;</p>
<!--#include virtual="/includes/footer.asp" -->
