<%
Option Explicit

Dim cn, strAddType, strHrefType, strOrderBy, strType
Dim blnNaughty
blnNaughty = False
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" --> 
<!--#include virtual="/includes/cocktail_functions.asp" -->
<%
set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB

strType=Request("type")

If strType <> "" Then
	If NOT IsNumeric(strType) Then
		Response.redirect("/")
	End if
	strHrefType = "&type="&Request.QueryString ("type")
	If strType= "1" Then
		strTitle = "All Cocktails"
		strAddType = " WHERE Status=1 And type=(type | 1)"
	Elseif strType= "2" Then
		strTitle = "All Shooters"
		strAddType = " WHERE Status=1 And type=(type | 2)"
	ElseIf strType= "4"  Then
		strTitle = "All Non-Alcoholic Cocktails"
		strAddType = " WHERE Status=1 AND type=(type | 4)"
	ElseIf strType= "8" Then
		strTitle = "All XXX Rated Drinks"
		strAddType = " WHERE Status=1 AND type=(type | 8)"
		blnNaughty = True
	End If 
Else
	strTitle = "All Drinks"
	strAddType = " WHERE Status=1"
	strHrefType = ""
End If

If Request("naughty") = "" AND strType <> "8" Then
	strAddType = strAddType & " AND type<>(type | 8)"
	blnNaughty = False
Else
	strHrefType = strHrefType & "&naughty=ON"
	blnNaughty = True
End If

If Request("userrecipes") = "" Then
	strAddType = strAddType & " AND Len(usr)<=0"
Else
	strHrefType = strHrefType & "&userrecipes=ON"
End If
%>
<!--#include virtual="/includes/header.asp" -->
<%
If Request("orderby") <> "" Then
	Session("orderby") = Request("orderby")
End If
If Session("orderby") <> "" Then
	If Session("orderby") = "name" Then
		strOrderBy = " ORDER BY name ASC"
	ElseIf Session("orderby") = "rate" then
		strOrderBy = " ORDER BY rate DESC, accessed DESC"
	Else
		strOrderBy = " ORDER BY " & Session("orderby") & " DESC, name ASC"
	End If
Else
	strOrderBy = " ORDER BY accessed DESC"
	Session("orderby") = "accessed"
End If

strSQL = "SELECT name, ID, type FROM cocktail" & strAddType & strOrderBy

writeCocktailList strSQL, rs, cn, strTitle, strHrefType

rs.Close
Set rs = Nothing
cn.Close
Set cn = Nothing
%>
<FORM name="order" action="?type="<%=strType%> method="GET">
  <div align="center">
  <table border="0" cellpadding="2" cellspacing="0" id="table1">
	<tr>
		<td valign="top"><B>Order By :</B></td>
		<td valign="top"><B>&nbsp;<INPUT type="radio" name="orderby" value="accessed" onclick="order.submit()" <%If Session("orderby")="accessed" then%>checked<%End If%> id="fp1" checked></B><LABEL for="fp1">Times Viewed</LABEL><B> 
  <INPUT type="radio" value="name" name="orderby" onclick="order.submit()"  <%If Session("orderby")="name" then%>checked<%End If%> id="fp2"></B><LABEL for="fp2">Name</LABEL><B> 
  <INPUT type="radio" name="orderby" value="rate" onclick="order.submit()" <%If Session("orderby")="rate" then%>checked<%End If%> id="fp3"></B><label for="fp3">Rating</label></td>
	</tr>
	<tr>
		<td valign="top"><b>Include:</b></td>
		<td valign="top">
  <P align="left">
	<input type="checkbox" name="userrecipes" value="ON" id="fp5" onclick="order.submit()" <%If Request("userrecipes")="ON" Then%> CHECKED<%End If%>><label for="fp5">User submitted recipes</label></P>
		</td>
	</tr>
	</table>
  </div>
  <INPUT type="hidden" name="type" value="<%=strType%>">
</FORM>
<!--#include virtual="/includes/footer.asp" -->