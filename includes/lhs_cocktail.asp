<%On Error Resume Next%>

<div class="menu-header"><a href="/db/Cocktails.asp"><img border="0" src="/images/side_menus/Drinks_right.gif" alt="Cocktails and shooter drinks"></a></div>

<div class="item">
<a class="linksin" title="Cocktail drink recipes" href="/cocktails/">Cocktail recipes</a></div>

<div class="item">
<a class="linksin" href="/cocktails/?type=2">Shots</a></div>

<div class="item">
<a class="linksin" href="/cocktails/?type=4">Non-alcoholic cocktails</a></div>

<div class="item">
<a class="linksin" href="/db/search/searchByIngredient.asp">Cocktails by ingredient</a></div>

<div class="item">
<a class="linksin" href="/db/viewCocktail.asp?ID=<%=getCOWID(dayNumber())%>">Drink of the day</a></div>

<div class="item">
<a class="linksin" href="/db/stats/toptenlatest.asp">Latest drinks</a></div>

<div class="item">
<a href="/db/random.asp" class="linksin">Random drink</a></div>

<div class="item">
<a href="/db/stats/default.asp" class="linksin">Top ten cocktails</a></div>




<div class="menu-header">
<a href="/db/member/loginOut.asp">
<img border="0" src="/images/side_menus/Members.gif" alt="Members ara"></a></div>

<%If Session("firstName") = "" Then%>
	<div class="item">
	<a href="/db/member/loginOut.asp" class="linksin">Log in...</a> / <a href="/db/member/createAccount.asp" class="linksin">Register</a></div>
<%Else%>
	<div class="item">
	<a href="/db/member/loginOut.asp" class="linksin"><%=Session("firstName")%>&#39;s account</a></div>
<%End If%>

<div class="item">
<a href="/db/member/userHotList.asp" class="linksin <%If Session("firstName") = "" Then%>disabled<%End If%>">Your favourite recipes</a></div>

<div class="item">
<a href="/db/member/selectIngredients.asp" class="linksin <%If Session("firstName") = "" Then%>disabled<%End If%>">Your bar ingredients</a></div>

<div class="item">
<a href="/db/member/whatCanIMake.asp" class="linksin <%If Session("firstName") = "" Then%>disabled<%End If%>">What you can make</a></div>




<div class="menu-header">
<a href="/shop/products/cocktail-equipment.asp">
<img border="0" src="/images/side_menus/Offer.gif" alt="Bar cocktail equipment and accessories"></a></div>

<div class="item">
<a href="/shop/basket.asp" class="linksin" style="text-decoration: underline;"><strong>View my basket</strong></a></div>
<!--#include virtual="/includes/shop/categoriesleft.asp" -->
<hr style="margin-top: 3px;">

<div class="item">
<a class="linksin" href="/shop/delivery.asp">Delivery prices</a></div>

<div class="item">
<a href="/shop/customerservices.asp" class="linksin">Customer services</a></div>
<img alt="Payment methods" src="/images/template/cards.gif" width="149" height="73">
