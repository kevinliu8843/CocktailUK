<%Dim strScriptName%>
<div class="lrnavp" align="left">
<a href="/db/Cocktails.asp">
<img border="0" src="/images/side_menus/Drinks_right.gif" alt="Cocktails and shooter drinks" width="150" height="42"></a><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
<a class="linksin" title="Cocktail drink recipes" href="/db/viewAllCocktails.asp?type=1">
Cocktail recipes</a><br>
<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
<a class="linksin" href="/db/viewAllCocktails.asp?type=2">Shooters</a><br>
<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
<a class="linksin" href="/db/viewAllCocktails.asp?type=8">Naughty XXX drinks</a><br>
<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
<a class="linksin" href="/db/viewAllCocktails.asp?type=4">Non-alcoholic cocktails</a><br>
<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
<a class="linksin" href="/db/viewCocktail.asp?ID=<%=getCOWID(dayNumber())%>">
Drink 
of the day</a><br>
<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
<a class="linksin" href="/db/stats/toptenlatest.asp">Latest drinks</a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
	<a href="/db/random.asp" class="linksin">Random drink</a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
	<a href="/db/member/submitCocktail.asp" class="linksin">Submit a drink</a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
	<a href="/db/stats/default.asp" class="linksin">Top ten</a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>


	<a href="/shop/products/cocktail-equipment.asp">
	<img border="0" src="/images/side_menus/Offer.gif" width="150" height="42" alt="Bar cocktail equipment and accessories"></a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
	&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">&nbsp;<a href="/shop/basket.asp" class="linksin" style="text-decoration: underline;"><strong>View my basket</strong></a><br>
	<!--#include virtual="/includes/shop/categoriesleft.asp" -->
	<hr style="margin-top: 3px;">
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
	&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">&nbsp;<a class="linksin" href="/shop/delivery.asp">Delivery prices</a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
	&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">&nbsp;<a href="/shop/customerservices.asp" class="linksin">Customer services</a><br>&nbsp;<br>
	<img alt="Payment methods" src="/images/template/cards.gif" width="149" height="73">

	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
	<a href="/db/member/loginOut.asp">
	<img border="0" src="/images/side_menus/Members.gif" width="150" height="42" alt="Members ara"></a><br>
&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8"> 
	<a href="/db/member/loginOut.asp" class="linksin"><%If Session("firstName") <> "" Then%><%=Session("firstName")%>&#39;s 
	members area<%else%>Log in<%End If%></a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
<%If Session("firstName") = "" Then%>
	&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
	<a href="/db/member/createAccount.asp" class="linksin">Register (free)</a><br>
<%Else%>
	&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
	<a href="/db/member/selectIngredients.asp" class="linksin">In my bar</a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
	&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
	<a href="/db/member/userHotList.asp" class="linksin">My favourites</a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
	&nbsp;<img border="0" src="/images/side_menus/smallarrow.gif" width="8" height="8">
	<a href="/db/member/whatCanIMake.asp" class="linksin">What can I make?</a><br>
	<img border="0" src="/images/pixel.gif" width="4" height="4"><br>
<%End If%>

</div>