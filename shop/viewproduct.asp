<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<%
Dim objProd, intID, strActKeywords, strDefaultKeywords, strAlsoBought, strComments, strCategory
intID = Request("ID")
If intID <> "" AND IsNumeric(intID) Then
	Set objProd = New CProduct

	If IsNumeric(Request("catID")) AND Request("catID") <> "" Then
		objProd.SetCategory(Request("catID"))
		Call objProd.GetCategoryName()
		strCategory = objProd.m_strCategoryName
		Call objProd.Reset()
	End If
	
	objProd.SetProductID(intID)
	strTopTitle = objProd.DisplayTopTitle
	strTopTitle = objProd.m_strProductName
	If strCategory <> "" Then
		strTopTitle = strTopTitle & " (From " & strCategory & " in the Cocktail : UK Bar Equipment Shop)"
	Else
		strTopTitle = strTopTitle & " - Cocktail : UK Bar Equipment Shop"
	End If
	strTitle = objProd.DisplayTitle
	call objProd.GetKeywords(strDefaultKeywords, strActKeywords)
	strMetaKeywords = objProd.m_strProductName  & ", " & strActKeywords
	strMetaDescription = objProd.m_strMetaDescription
	strAlsoBought = "" 'GetAlsoBought(intID, objProd.m_strProductName)
	%>
	<!--#include virtual="/includes/header.asp" -->
	<!--#include virtual="/includes/shop/header.asp" -->
    <h2><%=objProd.m_strProductName%></h2>
    <%objProd.DisplayProduct()%>
    <%=strAlsoBought%>    
	<%
	Set objProd = Nothing
Else
	Response.Redirect("/")
End If%>
<!--#include virtual="/includes/shop/footer.asp" -->
<!--#include virtual="/includes/footer.asp" -->