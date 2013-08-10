<%
Option Explicit
strTitle = "Thank You"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<H2>Thank you</H2>

<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">

<%Select case Request("type")%>
<%Case "address"%>
	<P>Your details have been sent to me and I will personally ensure that the pack is sent to you ASAP.</P>
	<p>Lee Tracey, Webmaster</p>
	<P>Return to the <A href="/">homepage</A></P>
<%Case "cocktail"%>
	<P>Your e-cocktail was delivered intact.</P>
	<P>Return to the <A href="/">homepage</A></P>
<%Case "details"%>
	<P>Your details have been sent to you and will arrive shortly.</P>
	<P>Return to the <A href="/">homepage</A></P>
	<P>Return to the <A href="/account/login.asp">login screen</A></P>
<%Case "friend"%>
	<P>An e-mail recommending cocktail.uk.com&nbsp; has been sent to your friend.</P>
	<P>Return to the <A href="/">homepage</A></P>
<%Case "page"%>
	<P>The email has been sent to your friend.
	<P>Return to the <A href="/">homepage</A></P>
<%Case else%>
	<P>Your comments have been sent to us.</P>
	<P>Return to the <A href="/">homepage</A></P>
<%End Select%>
    </td>
  </tr>
</table>

<!--#include virtual="/includes/footer.asp" -->