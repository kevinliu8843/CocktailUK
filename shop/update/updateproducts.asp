<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/admin_functions.asp" -->
<%
blnSilent = True
Call UpdateProductTables()
response.redirect("/admin/shop/diagnostic.asp")
%>
