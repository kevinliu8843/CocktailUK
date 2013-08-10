<% 
Option Explicit 
Dim i, objProduct, cn
strTitle = "Server Transfer Diagnostics"
On Error Resume Next
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<%
Dim strDSOut, strCUKOut, fso, f, objXmlHttpCat

If Request("categories") = "true" Then
	Call setupCategories(NULL)
End If

Set fso = Server.Createobject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(Server.MapPath("/shop/cuk_update.txt"))
strCUKOut = f.ReadAll()
%>
<h2>Shop diagnostics</h2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
<h3>Drinkstuff Output:<BR>
    </h3>
<%=Replace(strDSOut, VbCrLf, "<BR>")%>
<h3>Cocktail:UK Output:<BR>
    </h3>
<%=Replace(strCUKOut, VbCrLf, "<BR>")%>
<OL>
  <LI>
<P align="left"><B><A class="linksin" href="/shop/update/updateproducts.asp?Get=true&force=true">Update all shops products (including regeneration of data)</A></B></LI>
  <LI>
<P align="left"><B>
<A class="linksin" href="/shop/update/updateproducts.asp?force=true">Update Shops Products</A></B></LI>
  <LI>
<P align="left"><B><A class="linksin" href="?categories=true">Re-Create Categories</A></B></LI>
</OL>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->