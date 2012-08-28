<%
Option Explicit
strTitle = "Top 10 Latest Drinks"

Dim cn, iPageSize, FSO, name, FileExists
Dim iRecordsShown, strType
on error resume next
%>

<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->

<%
set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB

strSQL = "SELECT Top 10 name, ID, type, usr, dateadded FROM cocktail WHERE usr <> '' AND Status=1 ORDER BY dateadded DESC"

rs.Open strSQL, cn, 3, 3
iPageSize = 11
Set fso = CreateObject("Scripting.FileSystemObject")
name = Server.URLEncode(Replace( replaceStuffBack( rs("name") ), ",", "" ))

IF fso.FileExists(Server.MapPath("/images/cocktailThumbs/" & name & ".jpg" ) ) Then
	fileExists = True
Else
	fileExists = False
End If
%>
<h2>Top 10 latest drinks</h2>
  <TABLE border="0" cellpadding="0" cellspacing="20" width="100%">
    <TR>
      <TD width="10%">&nbsp;</TD>
      <TD align="center" width="80%" colspan="2" nowrap>
      <P>After some new, refreshing ideas? <BR>
      Then check out these latest submitted recipes...<BR>
      <A href="javascript:addToFavourites()">Click here and bookmark this page for the future</A></TD>
      <TD align="right" width="10%">&nbsp;</TD>
    </TR>

<%
iRecordsShown = 1
Do While NOT rs.EOF
	If iRecordsShown < iPageSize Then
		fileExists = False
	IF Int(rs("type")) AND 1 Then
		strType = "cocktail"
	Else
		strType = "shooter"
	End If
		%>

			<TR>
			  <TD width="10%">&nbsp;<FONT size="5" color="#FF0000"><%=iRecordsShown%></FONT></TD>
			  <TD align="right" width="40%" valign="bottom">&nbsp; <FONT size="1"><EM><A href="/db/viewCocktail.asp?ID=<%=rs("ID")%>"><STRONG><%=Capitalise(replaceStuffBack(rs("name"))) %></STRONG></A></EM><BR><%="Category: "&strtype%></FONT><BR><BR><B><FONT size=1>Added by <%=rs("usr")%> on <FONT size=1 color=gray><%=Day(rs("dateadded")) & " " & MonthName(Month(rs("dateadded")),True) & " " & Year(rs("dateadded"))%></FONT></FONT></TD>
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

	fileExists = False
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
			</FONT><FONT size="1"><EM><A href="/db/viewCocktail.asp?ID=<%=rs("ID")%>"><STRONG><%=Capitalise(replaceStuffBack(rs("name")))%></STRONG></A></EM><BR><%="Category: "&strtype%></FONT>&nbsp;&nbsp;<BR><BR><B><FONT size=1>Added by <%=rs("usr")%> on <FONT size=1 color=gray><%=Day(rs("dateadded")) & " " & MonthName(Month(rs("dateadded")),True) & " " & Year(rs("dateadded"))%></FONT></FONT></TD>
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

<P align="center">Feeling creative or just know a good recipe not on the site?<BR>
<A href="/db/member/submitcocktail.asp">The come and add your own recipe...</A></P>

<!--#include virtual="/includes/footer.asp" -->