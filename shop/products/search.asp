<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<%
Dim objProd, strSearch, strMainTitle
strSearch = Trim(Request("search"))
If Request("pagesize") <> "" Then
	Session("pagesize") = Request("pagesize")
End If
If strSearch <> "" Then
	Set objProd = New CProduct
	strSearch = Replace(strSearch, "'", "")
	objProd.SetSearchQuery(strSearch)
	strTopTitle = "Cocktail : UK > " & objProd.DisplayTopTitle
	strTitle = "Product " & objProd.DisplayTitle
	strMainTitle = "<H3>" & objProd.DisplayTitle & "</H3> "
	%>
	<!--#include virtual="/includes/header.asp" -->
	<!--#include virtual="/includes/shop/header.asp" -->
<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <TR>
    <TD valign="top" class="shopheadertitle">
	<H3><%=strMainTitle%></H3>
    </TD>
    <TD valign="top" align="right" class="shoplinebgonly" nowrap><FORM method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%><%If Request.QueryString <> "" Then%>?<%=Request.QueryString%><%End If%>">Show <SELECT size="1" name="pagesize" style="font-size: 8pt" onChange="this.form.submit()">
      <OPTION value="5" <%If Session("pagesize") = "5" Then%>selected<%End If%>>5</OPTION>
      <OPTION value="10" <%If Session("pagesize") = "10" OR Session("pagesize") = "" Then%>selected<%End If%>>10</OPTION>
      <OPTION value="20" <%If Session("pagesize") = "20" Then%>selected<%End If%>>20</OPTION>
      <OPTION value="999" <%If Session("pagesize") = "999" Then%>selected<%End If%>>All</OPTION>
      </SELECT> products</FORM>
    </TD>
  </TR>
</TABLE>
<FIELDSET class="fieldsetshop">
		<!--Display the products-->
		<%objProd.DisplayProducts%>
		<!--End products-->
		<%Set objProd = Nothing%>
  <P></P>
</FIELDSET>
<%
	'Response.write("<BR><CENTER><IFRAME frameborder=""0"" scrolling=""no"" id=""s0"" name=""s0"" align=absmiddle border=0 height=56 width=340 src=""/db/search/ask_jeeves.asp?full=true&searchfor="&Server.URLEncode(Request("Search"))&"""></IFRAME></CENTER>")
	'If strSearch <> "" Then
	'	Response.write("<SCRIPT LANGUAGE=""VBScript"" SRC=""/sitesearch/srch.vbs""></SCRIPT>")
	'	Response.write("<SCRIPT LANGUAGE=""JavaScript"">window.status='Cocktail : UK product search for "&Server.HTMLEncode(strSearch)&"'</SCRIPT>")
	'	please_encrypt("<IFRAME id=""s1"" name=""s1"" align=absmiddle border=1 height=0 width=0 src=""/db/search/ask_jeeves.asp?searchfor="&Server.URLEncode(strSearch)&"""></IFRAME>")
	'	please_encrypt("<IFRAME id=""s2"" name=""s2"" align=absmiddle border=1 height=0 width=0 src=""/includes/lycos.asp?searchfor="&Server.URLEncode(strSearch)&"""></IFRAME>")
	'End If
%><%Else
	Response.Redirect("/shop")
End If%><!--#include virtual="/includes/shop/footer.asp" --><!--#include virtual="/includes/footer.asp" -->