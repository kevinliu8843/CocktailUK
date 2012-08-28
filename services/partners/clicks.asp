<%
Option Explicit
strTitle="Click Counter"
Dim cn, intPartner, intClicks, strURL
%>

<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<%
intPartner = Request("id")
set cn = Server.CreateObject("adodb.connection")
cn.Open strDB
Set rs = cn.Execute("SELECT clicks, url from counter WHERE partner=" & intPartner)
if NOT rs.EOF then
	intClicks = rs("clicks")
	strURL = rs("url")
%>
<H2>Clickthroughs</h2>
<h4>Creative : "<%=strURL%>"</H4>
<H4>Total clicks since launch : <%=intClicks %></H4>
<%
else
%>
<H2>Clickthroughs</h2>
<h4>No such creative</H4>
<%End If%>
<!--#include virtual="/includes/footer.asp" -->
