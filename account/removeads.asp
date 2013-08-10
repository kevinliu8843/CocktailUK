<%
Option Explicit
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
strTitle = "How to remove the adverts..."
%>
<!--#include virtual="/includes/header.asp" -->
<SCRIPT Language="JavaScript">
function addHomeToFavourites(){
	if (document.all){
		window.external.AddFavorite('http://www.cocktail.uk.com/','Cocktail : UK - classic cocktails made easy.')
	}
	else{
		alert("With your current browser we can not add this page to your favourites automatically. Please set manually.")
	}
}
</SCRIPT>
<H2>How to remove the adverts.</H2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
<P align="left">In order to remove the adverts you see at the top of the screen, all you need to do is:</P>
<OL>
  <LI>
<P align="left"><B><A href="javascript:addHomeToFavourites()">Bookmark the site</A></B><BR>
So you remember us and come back again and again (but only if you like the site!)<BR>
<BR>
AND<BR>
&nbsp;</LI>
<LI>
<P align="left"><B><A href="/account/createaccount.asp">Register (it's free!)</A></B><BR>
So you can keep a record of what your favourite drinks are and what's in your drinks cabinet. And thus removing the adverts for every future visit!</LI>
</OL>
<P align="left">And that's it! The adverts will automatically be removed. Simple and free. Enjoy.<P align="center"><IMG border="1" src="../../images/adsremoved.gif" width="400" height="136" alt="Look - no adverts!!!"><BR>
<FONT size="1">Look - No adverts!!!</FONT></P>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->