<% 
Option Explicit 
Dim i, objProduct
strTitle = "Test random products"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<%
For i=0 To UBound(aryRandomProducts)
	Response.write "<HR>"
	Set objProduct = New CProduct
	objProduct.SetProductID(aryRandomProducts(i))
	objProduct.SetOnlyProduct()
	objProduct.DisplayProducts()
	Set objProduct=Nothing
Next
%><!--#include virtual="/includes/footer.asp" -->