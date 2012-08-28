<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/basketfunctions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<%
strTitle 	= "Create a brochure"

Dim aryCat, i, strText, cn

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open strDB

If Request("create") <> "true" Then
	strSQL = "SELECT name, ID from dscategory WHERE hidden=0 ORDER by catorder"
	rs.open strSQL, cn, 0, 3
	If NOT rs.EOF Then
		aryCat = rs.getrows()
	Else
		ReDim aryCat(0,0)
	End If
	rs.close
End If

cn.close
Set cn = nothing
Set rs = nothing

If Request("create") <> "true" Then
%>
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" -->
<h2>Download our 'live' brochure</h2>
<FORM method="POST" action="brochure.asp" target="_blank">
  <table border="0" cellpadding="5" style="border-collapse: collapse" width="100%" id="table1">
	<tr>
		<td>Here you can create your own personalised brochure, with only the products you are interested in...<P><B>What products do you want included in your brochure?</B></P>
  <TABLE border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="table2">
    <TR>
      <TD><INPUT type="radio" value="category" name="type" id="fp3" checked></TD>
      <TD><LABEL for="fp3">A Specific Category:</LABEL></TD>
      <TD><SELECT size="1" name="category">
      <%For i=0 to UBound(aryCat, 2)%>
      <OPTION value="<%=aryCat(1,i)%>"><%=aryCat(0,i)%></OPTION>
      <%Next%>
      </SELECT></TD>
    </TR>
    <TR>
      <TD><INPUT type="radio" name="type" value="search" id="fp2"></TD>
      <TD><LABEL for="fp2">Product Search:</LABEL></TD>
      <TD><INPUT type="text" name="search" size="20"></TD>
    </TR>
  </TABLE>
  <P align="center">
	<INPUT type="submit" value="Create Brochure &gt; &gt;" name="B1" class="button" ></P>
		</td>
	</tr>
	</table>
  <INPUT type="hidden" name="create" value="true">
</FORM>
<!--#include virtual="/includes/footer.asp" -->
<%Else%>
<HEAD>
<LINK rel="stylesheet" type="text/css" href="/style/style.css">
<TITLE>Cocktail : UK Brochure</TITLE>
</HEAD>
<BODY style="background-url:('/images/pixel.gif');">
<div align="center">
<table width="480" bgcolor="#FFFFFF"><tr><td>
<h3 align="center"><b><IMG border="0" src="/images/template/cuk_logo_banner.gif" align="absmiddle"> <font color="#808080"><i>Catalogue <%=Year(Now())%></i></font></b></h3>
</td></tr>
<tr><td>
<%
Dim objProd

Set objProd = New CProduct

Select Case Request("type")
	Case "category"
		objProd.SetCategory(Request("category"))
		strText = "Products from the category: " & objProd.DisplayTitle()
	Case "search"
		objProd.SetSearchQuery(Replace(Request("search"), "'", "''"))
		strText = "Here are your products containing the phrase: """ & Replace(Request("search"), "'", "''") & """"
End Select
objProd.SetPageSize(9999)
Response.write ("<P align=""center""><B>"&strText&"</B></P>")
objProd.DisplayProducts
Set objProd = Nothing
%>
</td></tr></table>
</div>
<SCRIPT Language="Javascript">
window.print();
</SCRIPT>
</BODY>
<%End If%>