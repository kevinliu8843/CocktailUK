<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-gb">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Find a drink</title>
<BASE TARGET="_main">
<SCRIPT>
function checkSearch()
{
	if ( document.search.searchField.value == " ) 
	{
		alert("Please enter a search query.")
		document.search.searchField.focus()
		return false
	}
	else
		return true
}
</SCRIPT>
</head>

<body style="background-color: #EFE3F0" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">

<table border="0" cellpadding="0" style="border-collapse: collapse" width="100%">
	<tr>
		<td bgcolor="#FFFFFF"><a href="http://www.cocktail.uk.com/">
<img border="0" src="../../../images/template/cuk_logo_banner.gif" width="300" height="80"></a></td>
	</tr>
</table>

<%Call DrawSearchCocktailArea()%>
<%If blnDoYahooThing Then%>
	<CENTER>
	<IFRAME width="100%" height="100" align="middle" frameborder="0" name="I1" scrolling="no" src="/includes/search_yahoo.asp"></IFRAME></CENTER>
<%End If%>
</body>

</html>