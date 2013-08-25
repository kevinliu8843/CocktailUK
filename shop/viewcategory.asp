<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/product.asp" -->
<%
Dim objProd, intCategory, strMainTitle
intCategory = Request("ID")
If Request("pagesize") <> "" Then
	Session("pagesize") = Request("pagesize")
End If
Set objProd = New CProduct
objProd.SetCategory(intCategory)
strTitleAppend = objProd.DisplayTopTitle
strTitle = objProd.DisplayTitle
strMainTitle = objProd.DisplayTitle
strMetaKeywords = objProd.GetCategoryKeywords
%>
<!--#include virtual="/includes/header.asp" -->
<H2 style="margin-bottom: 1em;"><%=strMainTitle%></h2>
<!--Display the products-->
<%objProd.DisplayProducts%>  
<!--End products-->
<%Set objProd = Nothing%>
<!--#include virtual="/includes/footer.asp" -->
