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

<h2>Cocktail recipes &amp; bar equipment store</h2>
<p><img border="0" src="../images/redmartini.jpg" width="125" height="244" align="right">
<p>Online for over 15 years, Cocktail : UK has the most comprehensive
<a href="/cocktails/">cocktail recipes</a> &amp; <a href="/cocktails/?type=2">shooter recipes</a> database 
(10,000+ and growing daily) plus
the biggest <a href="/shop/">bar equipment</a> shop online. Enjoy.</p>

<div>&nbsp;</div>

<div align="center"><a href="/cocktails/ingredients.asp">Find a 
drink</a> | <a href="/cocktails/">Cocktails</a> |
<a href="/cocktails/?type=2">Shooters</a> |
<a href="/cocktails/?type=4">Non-alcoholic cocktails</a> | 
<a href="shop/products/cocktail-equipment.asp">Bar equipment</a></div>

<H3 style="margin-top: 80px">Browse Our Cocktail Recipes</H3>
<TABLE border="0" cellpadding="5" cellspacing="0" id="table2" width="100%">
  <TR>
    <TD valign="top" align="left" colspan="2">
      <h4><A href="/cocktails/">View all 7000+ cocktails</A></h4>
      <div>Shows a list of all cocktails.</div>
      </TD>
  </TR>
  <TR>
    <TD valign="top" align="left"><h4><A href="/cocktails/basedon.asp?basedID=1">Vodka based cocktails</A></h4>
      <div>Shows a list of cocktails based on vodka.</div></TD>
    <TD valign="top" align="left">
      <h4><A href="/cocktails/basedon.asp?basedID=3">Gin based cocktails</A></h4>
      <div>Shows a list of cocktails based on gin.</TD>
  </TR>
  <TR>
    <TD valign="top" align="left">
      <h4><A href="/cocktails/basedon.asp?basedID=2">Brandy based cocktails</A></h4>
      <div>Shows a list of cocktails based on brandy.</TD>
    <TD valign="top" align="left"><h4><A href="/cocktails/basedon.asp?basedID=5">Rum based cocktails</A></h4>
      <div>Shows a list of cocktails based on rum.</TD>
  </TR>
  <TR>
    <TD valign="top" align="left">
      <h4><A href="/cocktails/basedon.asp?basedID=4">Whisky based cocktails</A></h4>
      <div>Shows a list of cocktails based on whisky.</TD>
    <TD valign="top" align="left">
      <h4><A href="/cocktails/basedon.asp?basedID=8">Tequila based cocktails</A></h4>
      <div>Shows a list of cocktails based on tequila.</TD>
  </TR></TABLE>
<!--#include virtual="/includes/footer.asp" -->