<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<%
Dim objProd, intCategory, strMainTitle
intCategory = 28
Set objProd = New CProduct
objProd.SetCategory(intCategory)
strTitleAppend = objProd.DisplayTopTitle
strTitle = objProd.DisplayTitle
strMainTitle = "<H3>" & objProd.DisplayTitle & "</H3>"
strKeywords = objProd.GetCategoryKeywords
%>
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" -->
<FIELDSET class="fieldsetshop">
  <LEGEND><B><%=strMainTitle%></B></LEGEND>
	<!--Display the products-->
	<%objProd.DisplayProducts%>
	<!--End products-->
  <P></P>
</FIELDSET><!--#include virtual="/includes/shop/footer.asp" --><!--#include virtual="/includes/footer.asp" -->
