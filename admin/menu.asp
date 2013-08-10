<%option explicit%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/admin_functions.asp" -->
<!--#include virtual="/includes/functions.asp" --><%
'On Error Resume Next
Dim dArrSales(10), iArrVolume(10), dArrProfits(10)
Dim dArrDSSales(10), iArrDSVolume(10), dArrDSProfits(10)
Dim dArrBARSales(10), iArrBARVolume(10), dArrBARProfits(10)
Dim dArrBARAffSales(10), iArrBARAffVolume(10), dArrBARAffProfits(10)
Dim dArrCEAffSales(10), iArrCEAffVolume(10), dArrCEAffProfits(10)
Dim intNewdrinks, intVisits, cn, intForumsReindex, intCocktailsReindex
Dim dteLastUpdated, intNewReviews, intNewGames, intShopReindex, intLinks, int404
Dim strWhen, dblDBsize, intProducts, intCategories, blnShowPC, dblDBSizeNew
Dim blnSilent

On Error Resume Next

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

cn.Open strDB

intNewdrinks = 0
intProducts = 0
intCategories = 0

strSQL = "SELECT count(*) from cocktailreview WHERE status=0"
rs.Open strSQL, cn
intNewReviews = rs(0)
rs.Close

strSQL = "SELECT count(*) from drink_desc WHERE status=0"
rs.Open strSQL, cn
intNewReviews = intNewReviews +rs(0)
rs.Close

strSQL = "SELECT count(*) from drinkinggame WHERE status=0"
rs.Open strSQL, cn
intNewGames= rs(0)
rs.Close

Call setupCategories(NULL)
Call CreatePrettyURLFiles(cn, rs)

Session("admin") = True

blnSilent = False
%> <HTML>

<HEAD>
<META http-equiv="Content-Language" content="en-gb">
<META http-equiv="Content-Type" content="text/html; charset=windows-1252">
<TITLE>Admin Menu</TITLE>
<BASE target="main">
<STYLE>
a{ text-decoration: none; }
</STYLE>
</HEAD>

<BODY topmargin="5" leftmargin="5" link="#000000" vlink="#000000" alink="#AA0000">
<SPAN class="linksin">

 <%If intNewdrinks>0 OR intNewReviews>0 OR intNewGames>0 Then%>
    <P><font face="Verdana"><font size="1">&nbsp;</font><B><font size="1" color="#612B83">To 
	Do List</font></B><font size="1"><BR>
 <%End If%>
 <%if intNewdrinks>0 Then%> &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="submitCocktails/default.asp">
	 <span style="text-decoration: none">Add Users' Drinks (<%=intNewdrinks%> new)</span></A></B><BR>
 <%End If%>
 <%if intNewReviews>0 Then%> &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="review/default.asp">
	 <span style="text-decoration: none">Add Users' Reviews (<%=intNewReviews%> new)</span></A></B><BR>
 <%End If%>
 <%if intNewGames>0 Then%> &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="game/default.asp">
	 <span style="text-decoration: none">Add Users' Games (<%=intNewGames%> new)</span></A></B><BR>
 <%End If%>
 </font>
 <P><font face="Verdana"><font size="1">&nbsp;</font><font color="#612B83" size="1"><B>Managers</B></font><font size="1"><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/admin/db/convert.asp">
 <span style="text-decoration: none">Restore Non-Alcoholic Drinks</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/admin/cocktaileditor">
 <span style="text-decoration: none">Edit Drinks</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="ingredients/modify_ingredients.asp">
 <span style="text-decoration: none">Edit Ingredients</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="ingredients/modify_measures.asp">
 <span style="text-decoration: none">Edit Measures</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/db/member/submitCocktail.asp">
 <span style="text-decoration: none">Add Drink Recipes</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="createHeaderAndFooter.asp">
 <span style="text-decoration: none">Create Header and Footer</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font></P>

    &nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://www.affiliatewindow.com/">
 <span style="text-decoration: none">AffiliateWindow</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://adstats.adviva.net/publisher/index.php">
 <span style="text-decoration: none">Adviva</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="https://euspecifics.specificmedia.com/">
 <span style="text-decoration: none">Specific Media</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://www.google.com/adsense">
 <span style="text-decoration: none">Google Adsense</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> </font></font></P>
</HTML>
<%
cn.Close
Set cn = Nothing
Set rs = Nothing
%>