<%
Option Explicit
strTitle = "Top Ten Most Popular"

Dim cn, iPageSize, fileExists, FSO, iRecordsShown, name, strType
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<%
set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB

strSQL = "SELECT Top 10 name, ID, type, accessed FROM cocktail ORDER BY accessed DESC"

iPageSize = 11
rs.pagesize = iPageSize

rs.Open strSQL, cn, 3, 3
Set fso = CreateObject("Scripting.FileSystemObject")
name = Replace( replaceStuffBack( rs("name") ), ",", "" )

IF fso.FileExists(Server.MapPath("/images/cocktailThumbs/"& name &".jpg" ) ) Then
	fileExists = True
Else
	fileExists = False
End If
%>
<h2>Top 10 Most Popular Cocktail Recipes</h2>
  <TABLE border="0" cellpadding="0" cellspacing="20" width="100%">
    <%
iRecordsShown = 1
Do While NOT rs.EOF
	If iRecordsShown < iPageSize Then
	name = Replace( replaceStuffBack( rs("name") ), ",", "" )

	IF fso.FileExists(Server.MapPath("/images/cocktailThumbs/"& name &".jpg" ) ) Then
		fileExists = True
	Else
		fileExists = False
	End If
	IF Int(rs("type")) AND 1 Then
		strType = "cocktail"
	Else
		strType = "shooter"
	End If
%>
		<TR>
		  <TD width="10%">&nbsp;<FONT size="5" color="#FF0000"><%=iRecordsShown%></FONT></TD>
		  <TD align="right" width="40%" valign="bottom">&nbsp; <FONT size="1"><EM><A href="/db/viewCocktail.asp?ID=<%=rs("ID")%>"><STRONG><%=replaceStuffBack(rs("name"))%></STRONG></A></EM><BR><%="Category: "&strtype%><BR>This <%=strType%> has been viewed <%=rs("accessed")%> times</FONT></TD>
		  <TD align="left" width="40%" valign="top">
			<P align="left"><FONT size="3"><A href="/db/viewCocktail.asp?ID=<%=rs("ID")%>"><IMG src="/images/cocktailThumbs/<%IF fileExists Then%><%=Replace( replaceStuffBack( rs("name") ), " ", "%20" )%><%Else%><%=strType%><%End If%>.jpg" border=0 height="100"></a><BR>
			</FONT></TD>
		  <TD align="right" width="10%"></TD>
		</TR>
<%
		' Increment the number of records we've shown
		iRecordsShown = iRecordsShown + 1
	End If

	If iRecordsShown < iPageSize AND ( NOT rs.EOF ) Then
		rs.MoveNext
		If NOT rs.EOF Then
	name = Replace( replaceStuffBack( rs("name") ), ",", "" )

	IF fso.FileExists(Server.MapPath("/images/cocktailThumbs/"& name &".jpg" ) ) Then
		fileExists = True
	Else
		fileExists = False
	End If
	IF Int(rs("type")) AND 1 Then
		strType = "cocktail"
	Else
		strType = "shooter"
	End If
	%>    
		<TR>
		  <TD width="10%"></TD>
		  <TD align="right" width="80%" valign="bottom" colspan="2">
			<HR size="1" color="#000000">
		  </TD>
		  <TD align="right" width="10%"></TD>
		</TR>
    
		<TR>
		  <TD width="10%">
			<FONT size="5" color="#FF0000"><%=iRecordsShown%></FONT>&nbsp;</TD>
		  <TD align="right" width="40%" valign="top"><FONT size="3"><A href="/db/viewCocktail.asp?ID=<%=rs("ID")%>"><IMG src="/images/cocktailThumbs/<%IF fileExists Then%><%=Replace( replaceStuffBack( rs("name") ), " ", "%20" )%><%Else%><%=strType%><%End If%>.jpg" border=0 height="100"></a></FONT></TD>
		  <TD width="40%" valign="bottom"><FONT size="3"><BR>
			</FONT><FONT size="1"><EM><A href="/db/viewCocktail.asp?ID=<%=rs("ID")%>"><STRONG><%=replaceStuffBack(rs("name"))%></STRONG></A></EM><BR><%="Category: "&strtype%><BR>This <%=strType%> has been viewed <%=rs("accessed")%> times</FONT></TD>
		  <TD width="10%">&nbsp;</TD>
		</TR>
		<TR>
		  <TD width="10%">
		  </TD>
		  <TD align="right" width="80%" nowrap valign="top" colspan="2">
			<HR size="1" color="#000000">
		  </TD>
		  <TD width="10%"></TD>
		</TR>
	<%
		' Increment the number of records we've shown
		iRecordsShown = iRecordsShown + 1
		End If
	End If

	If NOT rs.EOF Then
		rs.MoveNext
	End If
LOOP

rs.Close
Set rs = Nothing
cn.Close
Set cn = Nothing
%>

	 </TABLE>

<!--#include virtual="/includes/footer.asp" -->