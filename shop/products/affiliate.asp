<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<%
Dim objProd, strAffiliate
If Request("pagesize") <> "" Then
	Session("pagesize") = Request("pagesize")
End If
strAffiliate = "10724"
If strAffiliate = "" Then
	strAffiliate = Session("affID")
End If
If strAffiliate <> "" Then
	Set objProd = New CProduct
	Call objProd.SetAffiliate(strAffiliate)
	strTopTitle = objProd.DisplayTopTitle
	strTitle = objProd.DisplayTitle
	%>
	<!--#include virtual="/includes/header.asp" -->
	<!--#include virtual="/includes/shop/header.asp" -->
	<%If Request("shophome") = "true" Then%> 
		<H3>Welcome to the Cocktail : UK shop</H3>
		<P></P>
		<P>We are proud to open Cocktail : UK's new equipment shop. Now you can buy all your favourite products and browse all of your favourite recipes all in one place. </P>
	<%End If%>

<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <TR>
    <TD valign="top" class="shopheadertitle"> 
	<H3>Our current special offers</H3>
    </TD>
    <TD valign="top" align="right" class="shoplinebgonly" nowrap><FORM method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%><%If Request.QueryString <> "" Then%>?<%=Request.QueryString%><%End If%>">Show <SELECT size="1" name="pagesize" style="font-size: 8pt" onChange="this.form.submit()">
      <OPTION value="5" <%If Session("pagesize") = "5" Then%>selected<%End If%>>5</OPTION>
      <OPTION value="10" <%If Session("pagesize") = "10" OR Session("pagesize") = "" Then%>selected<%End If%>>10</OPTION>
      <OPTION value="20" <%If Session("pagesize") = "20" Then%>selected<%End If%>>20</OPTION>
      <OPTION value="999" <%If Session("pagesize") = "999" Then%>selected<%End If%>>All</OPTION>
      </SELECT> products
    </FORM>
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
Else
	Response.Redirect("/shop/")
End If
%><!--#include virtual="/includes/shop/footer.asp" --><!--#include virtual="/includes/footer.asp" -->