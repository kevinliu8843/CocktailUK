<%
Option Explicit
strTitle="E-News Subscription"

Dim cn, email
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<%
email = Request("email")

If email = "" Then
%>
<H2>Unsubscribe</H2>
<P>Please enter you e-mail address used when signing up to Cocktail : UK:&nbsp;
<form method="POST" action="unsubscribe.asp" webbot-action="--WEBBOT-SELF--">
  <p><input type="text" name="email" size="20"><input type="submit" value="Submit" name="B1"><input type="reset" value="Reset" name="B2"></p>
</form>
<P>

<%Else%>
<%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open(strDB)
strSQL = "UPDATE usr SET news=false WHERE email='" & Replace(strIntoDB(email), "'", "") & "'"
Set rs = cn.Execute( strSQL )
Set rs = Nothing
cn.Close
Set cn = Nothing
%>
<H2>Unsubscribe Confirmed</H2>
<P>You (<%=email%>) have now been unsubscribed from the Cocktail : UK E-Zine.
<P>Thank you for visiting Cocktail : UK
<P>Lee Tracey, Cocktail : UK
<%End If%><!--#include virtual="/includes/footer.asp" -->