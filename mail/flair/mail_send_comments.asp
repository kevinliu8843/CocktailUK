<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
strBody = "<HTML><HEAD><TITLE>Enquiry from Cocktail : UK (Courses)</TITLE><STYLE>body{font-family: Verdana, Arial, Geneva, sans-serif; font-size: 12px; color: #333366;}</STYLE></HEAD><BODY bgcolor=#ffffff><H2>Enquiry from Cocktail : UK (Courses)</H2><P><B>"&Request("forename")&"</B> (<a href=""mailto:"&Request("email")&""">"&Request("email")&"</A> Phone no: "&Request("phone")&") has made a query : <P>"&Replace(Request("S1"),VbCrLf, "<BR>")&"</BODY></HTML>"
strSubject = "Enquiry from Cocktail : UK (Courses)"
call SendEmail(Request("email"), "theresa@shaker-uk.com", "", "", strSubject, strBody, true, "")
Response.Redirect ("thankYou.asp")
%>