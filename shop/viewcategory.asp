<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<%
Dim objProd, intCategory, strMainTitle, strNoScript, strBannerTargetCat, strBannerURLCat, strBannerImageSrcCat
intCategory = Request("ID")
If Request("pagesize") <> "" Then
	Session("pagesize") = Request("pagesize")
End If
Set objProd = New CProduct
objProd.SetCategory(intCategory)
strTitleAppend = objProd.DisplayTopTitle
strTitle = objProd.DisplayTitle
strMainTitle = "<H3>" & objProd.DisplayTitle & "</H3>"
strMetaKeywords = objProd.GetCategoryKeywords
strNoScript = objProd.GetNoScript
Call objProd.GetBanner(strBannerTargetCat, strBannerURLCat, strBannerImageSrcCat)
%>
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" -->
<H2><%=strMainTitle%></h2>
<!--Display the products-->
<%objProd.DisplayProducts%>  
<!--End products-->
<%Set objProd = Nothing%>
<!--#include virtual="/includes/shop/footer.asp" -->
<!--#include virtual="/includes/footer.asp" -->
