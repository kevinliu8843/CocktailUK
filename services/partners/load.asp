<%
Option Explicit

Dim intPartner, cn, intClicks, strURL
%>
<!--#include virtual="/includes/variables.asp" -->
<%
intPartner = Request("id")
set cn = Server.CreateObject("ADODB.Connection")
cn.Open strDB
Set rs = cn.Execute("SELECT clicks, url from counter WHERE partner="&intPartner)
intClicks = rs("clicks")
strURL = rs("url")
Set rs = cn.Execute("UPDATE counter SET clicks=" & intClicks + 1 & " WHERE partner=" & intPartner)
response.Redirect(strURL)
%>