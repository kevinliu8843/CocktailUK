<% 
Option Explicit 

Dim cn, intUsers, intRate, strColourOver, strcolourOut, dblCocktails
Dim objProduct, objForumThread, bResult

strMetaDescription = "Buy bar cocktail equipment, Take Cocktail courses and buy poker chips, cocktails, Vodka shooters and non alcoholic cocktails from the experts."
strMetaKeywords = "cocktail uk drink cocktail recipes cocktails non alcoholic vodka shooter bar equipment poker chips uk courses and recipe"
strMetaTitle = "Cocktail UK - Cocktail Recipes & Cocktails / Shooters. Full Bar Equipment Store, How To Make Drinks Recipes."
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<script>
function checkSearch(){
	if ( document.search.searchField.value == "" ) {
		alert("Please enter a search query.")
		document.search.searchField.focus()
		return false
	}
	else
		return true
}
function clearField(){
	var strSearch = document.search.searchField.value
}

function changeColour(objTable, strColour){
	if (document.all){
		objTable.style.backgroundColor = strColour
	}
}
</script>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber14">
  <tr>
    <td width="100%" colspan="2" bgcolor="#FFFFFF">
    <table border="0" cellpadding="0" style="border-collapse: collapse; border-bottom: 1px solid #636388;" width="100%">
		<tr>
			<td width="180" valign="top"><a href="/db/cocktails.asp">
			<img border="0" src="images/homepage/martini.jpg" width="180" height="240" longdesc="View all our drink recipes"></a></td>
			<td valign="top">
			<p align="center">
			<img border="0" src="images/welcomeTo.gif" width="120" height="28"><br>
			<img border="0" src="images/cuk_07.gif" width="210" height="32"></p>
			<p align="left">Online for over 10 years, Cocktail : UK has the most comprehensive
			<a href="/db/cocktails.asp">cocktail recipes</a> &amp; <a href="/db/viewAllCocktails.asp?type=2">shooter recipes</a> database 
			(10,000+ and growing daily) plus
			the biggest
			<a href="/shop/">bar equipment</a> shop online. Enjoy.<p align="center">
			<img border="0" src="images/dotted_line_horizontal.gif" width="162" height="1"><br>
&nbsp;<div align="center"><a href="/sitesearch/">Find a 
			drink</a> | <a href="db/viewAllCocktails.asp?type=1">Cocktails</a> |
			<a href="db/viewAllCocktails.asp?type=2">Shooters</a> |
			<a href="db/viewAllCocktails.asp?type=4">Non-alcoholic</a> <br>
			<a href="shop/products/cocktail-equipment.asp">Bar equipment</a> |
				<a href="/db/member/submitCocktail.asp">Submit a recipe</a></div>
			</td>
		</tr>
	</table>
    </td>
  </tr>
</table>
<H2>Cocktails</H2>
<table border="0" cellpadding="5" style="border-collapse: collapse" width="100%" id="table1">
  <tr>
    <td><img border="0" src="../images/redmartini.jpg" width="125" height="244"></td>
    <td>
<TABLE border="0" cellpadding="5" cellspacing="0" id="table2">
  <TR>
    <TD valign="top" align="left" colspan="3">
      <P align="left"><A href="/db/viewAllCocktails.asp?type=1"><B>View
      all cocktails<BR>
      </B></A>Shows a list of all cocktails.
      </TD>
  </TR>
  <TR>
    <TD valign="top" align="left"><A href="/based_cocktails/vodka.asp"><B>Vodka
      based<BR>
      </B></A>Shows a list of cocktails based on vodka.</TD>
    <TD valign="top" align="left">&nbsp;</TD>
    <TD valign="top" align="left">
      <P align="left"><A href="/based_cocktails/gin.asp"><B>Gin based<BR>
      </B></A>Shows a list of cocktails based on gin.</TD>
  </TR>
  <TR>
    <TD valign="top" align="left">
      <P align="left"><A href="/based_cocktails/brandy.asp"><B>Brandy
      based<BR>
      </B></A>Shows a list of cocktails based on brandy.</TD>
    <TD valign="top" align="left">&nbsp;</TD>
    <TD valign="top" align="left"><A href="/based_cocktails/rum.asp"><B>Rum
      based<BR>
      </B></A>Shows a list of cocktails based on rum.</TD>
  </TR>
  <TR>
    <TD valign="top" align="left">
      <P align="left"><A href="/based_cocktails/whisky.asp"><B>Whisky
      based<BR>
      </B></A>Shows a list of cocktails based on whisky.</TD>
    <TD valign="top" align="left">&nbsp;</TD>
    <TD valign="top" align="left">
      <P align="left"><A href="/based_cocktails/tequila.asp"><B>Tequila
      based<BR>
      </B></A>Shows a list of cocktails based on tequila.</TD>
  </TR>
</TABLE>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->