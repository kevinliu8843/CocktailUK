<%Option Explicit%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
strtitle="All drinking games"
Dim i, cn
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open strDB
strSQL = "SELECT count(*) from drinkinggame where status=1"
set rs = cn.execute(strSQL)
%>
<!--#include virtual="/includes/header.asp" -->
<H2>Drinking Games</H2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
  <tr>
    <td width="100%">Hi there, we hope <FONT face="Lazy dog">you</FONT> enjoy these drinking games...If you know one yourself that would make a good addition to this list, please <A href="submit_game.asp">submit it here</A>...
<% If request("gameadded") = "true" Then%>
	<P align="center"><I><FONT color="#FF0000">Thank you, your game has been submitted for review...</FONT></I>
<%End If%>
<P align="center">Please select a type of game from below
<P>
&nbsp;<DIV align="center">
  <CENTER><TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1">
  <TR>
    <TD nowrap valign="top" width="220">
	<P align="center"><A href="/shop/products/product.asp?ID=139"><B><IMG border="0" src="../../images/games/drinking/slambango.jpg" width="219" height="188"><BR>
    </B>Check out our slambango game for £14.99</A></TD>
    <TD valign="top" nowrap>
	<MENU>
	<LI><A HREF="view_games.asp">All games (<%=rs(0)%>)</A></LI>
	<%For i=0 to UBound(aryGames)%>
		<%strSQL = "SELECT count(*) from drinkinggame where type="&i&" and status=1"
		  set rs = cn.execute(strSQL)%>
		<LI><A HREF="view_games.asp?type=<%=i%>"><%=aryGames(i)%> (<%=rs(0)%>)</A></LI>
		<%Set rs=Nothing
	Next
	cn.close
	Set cn = Nothing
	%>
	</MENU>
    </TD>
  </TR>
</TABLE>
  </CENTER>
</DIV>
<P align="center"><A href="submit_game.asp"><IMG border="1" src="../../images/main_menus/addyourgame.gif" style="border-style: solid; border-color: #800080" width="150" height="23"></A></td>
    </tr>
  </table>
<!--#include virtual="/includes/footer.asp" -->