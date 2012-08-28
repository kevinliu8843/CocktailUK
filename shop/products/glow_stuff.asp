<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<%
Dim objProd, intCategory, strMainTitle
intCategory = 29
Set objProd = New CProduct
objProd.SetCategory(intCategory)
'strTopTitle = objProd.DisplayTopTitle
strTitle = objProd.DisplayTitle
strMainTitle = "<H3>" & objProd.DisplayTitle & "</H3>"
strKeywords = objProd.GetCategoryKeywords
%>
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" -->
<%=strMainTitle%>
<!--Display the products-->
<%objProd.DisplayProducts%>
<!--End products-->
<P>
<!--#include virtual="/includes/shop/footer.asp" -->
<!--#include virtual="/includes/footer.asp" -->
