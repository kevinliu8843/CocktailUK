<%Option Explicit%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<%
Dim i, cn, strType, intType, intID

intType = Request("type")

strTitle = "All Games"
strType = ""

If IsNumeric(intType) AND intType<>"" Then
	If Int(intType) <= UBound(aryGames)  then
		strTitle = aryGames(intType) & " Games"
		strType = " AND type=" & intType
	End If
End If
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open strDB
strSQL = "SELECT title, ID, rating from drinkinggame where status=1" & strIntoDB(strType) & " ORDER BY rating desc, title"
set rs = cn.execute(strSQL)
%>
<!--#include virtual="/includes/header.asp" -->
<H2><%=strTitle%></H2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%"><B><A href="javascript:history.go(-1)">&lt; &lt; Back</A></B><%If NOT rs.EOF Then%>
	<P>Please select a game from below<MENU>
	<%While NOT rs.EOF%>
		<LI><A HREF="view_game.asp?ID=<%=rs("ID")%>" border="0"><%call displayRatingGraphOnly( rs("rating") )%></A> <A HREF="view_game.asp?ID=<%=rs("ID")%>"><%=rs("title")%></A></LI>
		<%rs.moveNext
	Wend%>
<%Else%>
<P>Sorry, there are no <%=LCase(strTitle)%> at the moment. Do you know any good ones? <A href="submit_game.asp">Submit them here</A>...
<%End If%>
<%
cn.close
Set cn = Nothing
%>
<P align="center"><A href="submit_game.asp"><IMG border="1" src="../../images/main_menus/addyourgame.gif" style="border-style: solid; border-color: #800080" width="150" height="23"></A></P>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->