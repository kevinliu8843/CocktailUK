<%On Error Resume Next%>

<div class="menu-header"><a href="/cocktails/"><img border="0" src="/images/side_menus/Drinks_right.gif" alt="Cocktails and shooter drinks"></a></div>

<div class="item">
<a class="linksin" title="Cocktail drink recipes" href="/cocktails/">Cocktail recipes</a></div>

<div class="item">
<a class="linksin" href="/cocktails/?type=2">Shots</a></div>

<div class="item">
<a class="linksin" href="/cocktails/?type=4">Non-alcoholic cocktails</a></div>

<div class="item">
<a class="linksin" href="/cocktails/ingredients.asp">Cocktails by ingredient</a></div>

<div class="item">
<a href="/cocktails/random.asp" class="linksin">Random cocktail</a></div>

<div class="item">
<a href="/cocktails/top-ten/" class="linksin">Top ten cocktails</a></div>




<div class="menu-header">
<a href="/account/login.asp">
<img border="0" src="/images/side_menus/Members.gif" alt="Members ara"></a></div>

<%If Session("firstName") = "" Then%>
	<div class="item">
	<a href="/account/login.asp" class="linksin">Log in...</a> / <a href="/account/register.asp" class="linksin">Register</a></div>
<%Else%>
	<div class="item">
	<a href="/account/login.asp" class="linksin"><%=Session("firstName")%>&#39;s account</a></div>
<%End If%>

<div class="item">
<a href="/account/userHotList.asp" class="linksin <%If Session("firstName") = "" Then%>disabled<%End If%>">Your favourite recipes</a></div>

<div class="item">
<a href="/account/selectIngredients.asp" class="linksin <%If Session("firstName") = "" Then%>disabled<%End If%>">Your bar ingredients</a></div>

<div class="item">
<a href="/account/whatCanIMake.asp" class="linksin <%If Session("firstName") = "" Then%>disabled<%End If%>">What you can make</a></div>




<div class="menu-header">
<a href="/shop/products/cocktail-equipment.asp">
<img border="0" src="/images/side_menus/Offer.gif" alt="Bar cocktail equipment and accessories"></a></div>

<div class="item">
<!--#include virtual="/includes/shop/categoriesleft.asp" -->
<hr style="margin-top: 3px;">

<div class="item">
<a class="linksin" href="/shop/delivery.asp">Delivery prices</a></div>

<div style="margin-top: 10px;"><img alt="Payment methods" src="/images/template/cards.gif" width="149" height="73"></div>

<div style="margin-top: 10px;" class="item">
<a href="/services/privacy.asp" class="linksin">Privacy policy</a></div>
