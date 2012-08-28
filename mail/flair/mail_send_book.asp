<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
strBody = "<HTML><HEAD><TITLE>Course Booking Request from Cocktail : UK</TITLE><STYLE>body{font-family: Verdana, Arial, Geneva, sans-serif; font-size: 12px; color: #333366;}</STYLE></HEAD><BODY bgcolor=#ffffff><H2>Course Booking Request from Cocktail : UK</H2><P><B>"&Request("forename")&" "&Request("surname")&"</B> (<a href=""mailto:"&Request("email")&""">"&Request("email")&"</A> Phone no: "&Request("phone")&") has made a booking request : <P>"&Request("address")&"<BR>"&Request("postcode")&"<BR>"&Request("telephone")&"<P>For the course and dates:<BR>"&Request("course")&"<BR>"&Request("date")&"</BODY></HTML>"
strSubject = "Course Booking Request from cocktail : UK"
call SendEmail(Request("email"), "theresa@shaker-uk.com", "", "", strSubject, strBody, true, "")
Response.Redirect ("thankYouBook.asp")
%>