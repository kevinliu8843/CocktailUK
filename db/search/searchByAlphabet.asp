<%
Option Explicit

Dim cn, searchfor, strHrefType, char

'grab search field
searchfor = Request("letter")
If Len(searchfor) >1 Then
	searchfor = Left(searchfor, 1)
end if

If searchfor = "" then
	strTitle = "Search By Alphabet"
else	
	strTitle = "Drinks starting with " & searchfor 
end if
%>
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/header.asp" -->
<H2><%=strTitle %></H2>
<P align="center">
Choose the first letter of the drink
<P align="center"><b>
<% 
For char = Asc("A") To Asc("Z") %>
	<A href="searchByAlphabet.asp?letter=<%=LCase(Chr(char))%>"><%=Chr(char)%></A>&nbsp;
<% Next %>
</b>
</P>

<%If searchfor <> "" AND Len(searchfor) <= 1 Then%>
	<%
	set cn = Server.CreateObject("ADODB.Connection")
	Set rs = Server.CreateObject("ADODB.RecordSet")
	cn.Open strDB
	
	'Search for name only
	strSQL = "Select name, ID, type From cocktail WHERE Status=1 And name LIKE '" & strIntoDB(searchfor) & "%' ORDER by name"
	
	strHrefType = "&letter=" & searchfor
	writeCocktailList strSQL, rs, cn, "", strHrefType
	%>
<%Else%>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
	<p>&nbsp;</p>
<%End If%></B>
<%Call DrawSearchCocktailArea()%><!--#include virtual="/includes/footer.asp" -->