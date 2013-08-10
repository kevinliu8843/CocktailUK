﻿<%On Error Resume Next%>

<div class="menu-header"><a href="/db/Cocktails.asp"><img border="0" src="/images/side_menus/Drinks_right.gif" alt="Cocktails and shooter drinks"></a></div>

<div class="item">
<a class="linksin" title="Cocktail drink recipes" href="/db/viewAllCocktails.asp?type=1">Cocktail recipes</a></div>

<div class="item">
<a class="linksin" href="/db/viewAllCocktails.asp?type=2">Shots</a></div>

<div class="item">
<a class="linksin" href="/db/viewAllCocktails.asp?type=4">Non-alcoholic cocktails</a></div>

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
	<a href="/db/member/loginOut.asp" class="linksin">Log in...</a></div>

	<div class="item">
	<a href="/db/member/createAccount.asp" class="linksin">Register (<span style="color: #0990000;">free</psan>)</a></div>

	<div class="item">
	<a href="/db/member/userHotList.asp" class="linksin" disabled>Your favourite recipes</a></div>

	<div class="item">
	<a href="/db/member/selectIngredients.asp" class="linksin" disabled>Your bar ingredients</a></div>

	<div class="item">
	<a href="/db/member/whatCanIMake.asp" class="linksin" disabled>What you can make</a></div>

<%Else%>
	<div class="item">
	<a href="/db/member/loginOut.asp" class="linksin"><%If Session("firstName") <> "" Then%><%=Session("firstName")%>&#39;s 
	members area<%else%>Log in<%End If%></a></div>

	<div class="item">
	<a href="/db/member/userHotList.asp" class="linksin">Your favourite recipes</a></div>

	<div class="item">
	<a href="/db/member/selectIngredients.asp" class="linksin">Your bar ingredients</a></div>

	<div class="item">
	<a href="/db/member/whatCanIMake.asp" class="linksin">What you can make</a></div>
<%End If%>




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
