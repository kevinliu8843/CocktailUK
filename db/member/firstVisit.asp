<%
Option Explicit
strTitle="First Visit"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<%
Session("first") = "first"
%>

<H2>Welcome <%=Session("firstname")%> - first time login</H2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">As it is your first time logging in at Cocktail : UK the first thing we would like to say is welcome and we hope you enjoy your stay and that you come back and 
visit us often.<P> In your members section, you can have a favourites menu of the drinks you <B>love</B>, contribute to our forums and figure out which recipes you can make with the ingredients you already have at home. What could be easier?</P>

<P align="center"><A href="/db/member/loginOut.asp">Continue to the members section...</A></P>

    </td>
  </tr>
</table>

<!--#include virtual="/includes/footer.asp" -->