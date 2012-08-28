<!--#include virtual="/includes/variables.asp" -->
<%
If Request("competition") = "true" Then
	strExtra = "&page=/competition"
End if
If Request("c") = "1" Then
	response.redirect("http://www.drinkstuff.com/affiliate/cookie.asp?affID=10724")
	strExtra = "&blank=true"
End If
response.redirect("http://www.drinkstuff.com/products/affiliate.asp?affID=10724" & strExtra)
%>