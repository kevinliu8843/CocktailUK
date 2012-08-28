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
<table border="0" cellpadding="0" style="border-collapse: collapse">
  <tr>
    <td valign="top" width="100%">
<table border="0" cellpadding="0" style="border-collapse: collapse" id="table8">
  <tr>
    <td valign="top" width="50%" background="images/grad_write_purple.gif">
    <table border="0" cellpadding="0" style="border-collapse: collapse; border-bottom:1px solid #636388; " bordercolor="#993399" id="table9" height="165" width="0">
      <tr>
        <td height="10"><a href="/sitesearch/">
        <img border="0" src="images/main_menus/findthatdrink.gif" width="239" height="42"></a></td>
      </tr>
      <tr>
        <td background="images/grad_write_purple.gif">
        <table border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="table10" height="100%" cellpadding="2">
          <tr>
            <td width="100%">
            <p align="center">Ever had a wonderful cocktail and you just have no 
			idea how it was made? Then why not search for it and find out!...</p>
            <form method="POST" action="/sitesearch/default.asp" name="search2" onSubmit="return checkSearch()">
             <p align="center">
             <input type="text" name="searchField" size="24" style="border:1px solid #979797; width: 134; height: 19; text-align: left" class="shopoption"><input border="0" src="images/template/cuk_orange_btn_go.gif" name="I2" align="absmiddle" type="image"></p>
            </form>
            </td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
    <td valign="top" width="50%">
    <table border="0" cellpadding="0" style="border-collapse: collapse; border-bottom:1px solid #636388;" id="table11" height="165" width="0">
      <tr>
        <td height="10" background="images/main_menus/youringredients.gif"><a href="/db/member/selectIngredients.asp">
        <img border="0" src="images/pixel.gif" width="237" height="42"></a></td>
      </tr>
      <tr>
        <td background="images/grad_write_purple.gif" valign="top">
        <table border="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="table12" height="100%" cellpadding="2">
          <tr>
            <td width="100%" valign="top">
            <p align="center">Do you have some ingredients that you want to make 
			cocktails out of? Then tell us your ingredients and let us tell you 
			what drink recipes you can make!</p>
            <p align="center"><b><a href="/db/member/selectIngredients.asp">
			Enter your ingredients here...</a></b></p>
            </td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->