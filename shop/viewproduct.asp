<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/product.asp" -->
<%
Dim objProd, intID, strActKeywords, strDefaultKeywords, strAlsoBought, strComments, strCategory
intID = Request("ID")
If intID <> "" AND IsNumeric(intID) Then
	Set objProd = New CProduct
	objProd.SetProductID(intID)
	strTopTitle = objProd.m_strProductName & " - Cocktail : UK Bar Equipment Shop"
	strTitle = objProd.DisplayTitle
	call objProd.GetKeywords(strDefaultKeywords, strActKeywords)
	strMetaKeywords = objProd.m_strProductName  & ", " & strActKeywords
	strMetaDescription = objProd.m_strMetaDescription
	strAlsoBought = ""
	%>
	<!--#include virtual="/includes/header.asp" -->
    <h2><%=objProd.m_strProductName%></h2>
    <%objProd.DisplayProduct()%>
    <%=strAlsoBought%>    
	<%
	Set objProd = Nothing
Else
	Response.Redirect("/")
End If%>
<!--#include virtual="/includes/footer.asp" -->