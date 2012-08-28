<%
Option Explicit
Dim searchfor
If InStr(Request.ServerVariables("HTTP_REFERER"), "/sitesearch") > 0 OR InStr(Request.ServerVariables("HTTP_REFERER"), "/products/search.asp") > 0 OR InStr(Request.ServerVariables("HTTP_REFERER"), "/products/similar.asp") > 0 then
	searchfor = Request("searchfor")
Else
	searchfor = ""
End If
strTitle = "Search the web"
%>
<!--#include virtual="/includes/variables.asp" -->
<%If searchfor = "" Then%>
	<!--#include virtual="/includes/functions.asp" -->
	<!--#include virtual="/includes/header.asp" -->
	<H2>Search the web</H2>
<%ElseIf Request("full") <> "true"  Then%>
	<html>
	
	<head>
	<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
	<title>Search the web</title>
	</head>

	<h2>Search the web</h2>
	<p align="center"><b>Can't find what you are looking for?</b><CENTER>
<%Else%>
<BODY leftmargin=0 topmargin=0>
<%End If%>
<CENTER>
<!--#include virtual="/includes/search_ask_jeeves.asp" -->
<img src="http://wzeu.ask.com/i/i.gif?t=a&d=eu&s=uk&c=d&ti=2&ai=51223&l=dir&o=38881094&sv=0a652847&ip=c363a042&ord=6965783" border="0" width="1" height="1" />

</CENTER>
<%If searchfor = "" Then%>
	<!--#include virtual="/includes/footer.asp" -->
<%Else%>
	</body>

	</html>
<%End If%>