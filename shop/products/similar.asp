<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<!--#include virtual="/includes/shop/basketfunctions.asp" -->
<%
Dim objProd, strSearch, strMainTitle, strAlsoBought, intID
Dim aryExtraProds(3,4), intExtraProds, i

strSearch = Trim(Request("search"))
intID = Request("ID")
If intID = "" OR strSearch = "" OR NOT IsNumeric(intID) Then
	response.redirect("/shop/")
End If
strTopTitle = "Cocktail : UK Equipment Shop &gt; Similar products to " & strSearch 
strTitle = strSearch  & " (Similar products)"
strMainTitle = "<H3>Similar products to <FONT color=silver>" & strSearch & "</FONT></H3> "
call GetAlsoBoughtData(intID, aryExtraProds, intExtraProds)
strAlsoBought = GetAlsoBought(intID, strSearch)
 
Set objProd = New CProduct
%>
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" -->
	<H3><%=strMainTitle%></H3>
	<FIELDSET>
	<LEGEND><I><B>Finding products similar to <%=strSearch%></B></I></LEGEND>
	<%
	'Display the product in question
	call objProd.SetOnlyProduct()
	call objProd.SetProductID(intID)
	objProd.DisplayProducts
	%>
    <P></P>
	</FIELDSET>
	<%=strAlsoBought%>
    <P></P>
	
	<FIELDSET class="fieldsetshop">
	<%	
	'Display tghe 5 other similar products here
	For i=0 to intExtraProds-1
		call objProd.SetProductID(aryExtraProds(0, i))
		objProd.DisplayProducts
		response.write "<HR noshade color=""#336699"" size=""1"">"
	Next
	
    response.write "<TABLE border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""4"" background=""../images/grad_write_purple.gif"">"
    response.write "<TR>"
    response.write "<TD bgcolor=""#612b83"" height=""1""><FONT color=""#FFFFFF""><B>Searching for <I>"&strSearch&"</I> produced...</B></FONT></TD>"
    response.write "</TR></TABLE>"
                  
	objProd.m_blnOnlyProduct = False
	strSearch = Replace(strSearch, "'", "")
	objProd.SetSearchQuery(strSearch)
	%>
	<!--Display the products-->
	<%objProd.DisplayProducts%>
	<!--End products-->
	<%Set objProd = Nothing%>
  <P></P>
</FIELDSET>
<!--#include virtual="/includes/shop/footer.asp" --><!--#include virtual="/includes/footer.asp" -->