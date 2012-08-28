<%
Option Explicit

Dim cn, basedID, strHrefType, strSQLcocktail, basedOn, based, iPageCurrent, iPageSize, iPageCount
Dim FSO, name, FileExists, strType, strAddType, strOrderBy
%>
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" -->
 
<%
set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB

'--------------------------------------------------------
basedID = Request("basedID")

'Error trap
If basedID = "" OR NOT IsNumeric(basedID) Then basedID = 1

strHrefType = "&basedID="&strIntoDB(basedID)

strSQLcocktail = "SELECT name, ID, type FROM cocktail WHERE Status=1 And based=" & basedID

If CStr( basedID ) = "5" Then	'include dark rum based drinks too
	strSQL = "SELECT name, ID FROM ingredients WHERE ID=" & basedID & " OR ID=6 OR ID=7"
	strSQLcocktail = "SELECT name, ID, type FROM cocktail WHERE Status=1 And based IN (" & basedID & ",6,7)"
	based = "Rum"
Else
	strSQL = "SELECT name, ID FROM ingredients WHERE ID="&basedID
End If

If CStr( basedID ) = "8" Then	'include gold tequila based drinks too
	strSQL = "SELECT name, ID FROM ingredients WHERE ID=" & basedID & " OR ID=9"
	strSQLcocktail = "SELECT name, ID, type FROM cocktail WHERE Status=1 And based IN (" & basedID & ",9)"
	based = "Tequila"
Else
	strSQL = "SELECT name, ID FROM ingredients WHERE ID="&basedID
End If

If CStr( basedID ) = "4" Then	'include other whiskies too
	strSQL = "SELECT name, ID FROM ingredients WHERE ID IN (" & basedID & ",10,15)"
	strSQLcocktail = "SELECT name, ID, type FROM cocktail WHERE Status=1 And based IN (" & basedID & ",10,15)"
	based = "Whisky"
Else
	strSQL = "SELECT name, ID FROM ingredients WHERE ID="&basedID
End If

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

If Request("naughty") = "" Then
	strAddType = strAddType & " AND type<>(type | 8)"
Else
	strHrefType = strHrefType & "&naughty=ON"
End If

If Request("userrecipes") = "" Then
	strAddType = strAddType & " AND Len(usr)<=0"
Else
	strHrefType = strHrefType & "&userrecipes=ON"
End If

rs.Open strSQL, cn, 0, 3
If Not rs.EOF Then
	basedOn = rs("name")
End If
rs.Close
If NOT CStr( basedID ) = "5" Then
	based = capitalise( replaceStuffBack( basedOn ) )
End If

strTitle = "List of " & based & " based cocktails"

'--------------------------------------------------------
%>
<!--#include virtual="/includes/header.asp" -->
<%
Call writeCocktailList( strSQLcocktail & strAddType & strOrderBy, rs, cn, strTitle, strHrefType )
cn.Close
Set cn = Nothing
Set rs = Nothing
%>
<FORM name="order" action="viewBasedCocktails.asp" method="GET">
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
	<%If strType <> "8" Then%>
		<input type="checkbox" name="naughty" value="ON" id="fp4" onclick="order.submit()" <%If Request("naughty")="ON" Then%> CHECKED<%End If%>><label for="fp4">Naughty XXX drinks </label><img border="0" src="../images/s18.gif" width="25" height="25" align="absmiddle">
		<br>
	<%End If%>
		<input type="checkbox" name="userrecipes" value="ON" id="fp5" onclick="order.submit()" <%If Request("userrecipes")="ON" Then%> CHECKED<%End If%>><label for="fp5">User submitted recipes</label></P>
  		</td>
	</tr>
	</table>
  </div>
  <INPUT type="hidden" name="basedID" value="<%=basedID%>">
</FORM>
<%Call DrawSearchCocktailArea()%><!--#include virtual="/includes/footer.asp" -->