<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/admin_functions.asp" -->
<%
'This page gets called from "http://www.drinkstuff.com/admin/advanced/csv/export_to_cocktailuk.asp" but with no query string data
blnSilent = True
If Request("Get") = "true" Then
	Call GetProductTablesAndUpdate()
	Call ReindexSite("reindex shop", "product(s)", intShopReindex, blnSilent)
	If request("goto") = "" then
		Response.redirect("/default.asp")
	Else
		Response.redirect(request("goto"))
	End If
Else
	Call UpdateProductTables()
	Call ReindexSite("reindex shop", "product(s)", intShopReindex, blnSilent)
End If
%>
