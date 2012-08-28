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

blnShowPC = (Request("percentages") = "true")

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

cn.Open strDBMod
strSQL = "SELECT visitors FROM counter WHERE year= " & Year(Now) & " AND month=" & Month(Now)
rs.Open strSQL, cn
If NOT rs.EOF Then
 intVisits  = rs("visitors")
Else
 intVisits  = 0
End If
rs.Close

strSQL = "SELECT COUNT(*) FROM ForumMessages WHERE ReIndex=1"
rs.Open strSQL, cn
intForumsReindex = rs(0)
rs.Close

strSQL = "SELECT count(*) from Cocktail WHERE Status=0"
rs.Open strSQL, cn
intNewdrinks = rs(0)
rs.Close

strSQL = "SELECT count(*) from Cocktail WHERE ReIndex=1"
rs.Open strSQL, cn
intCocktailsReindex = rs(0)
rs.Close

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

strSQL = "SELECT count(*) from dsproduct WHERE status=1 AND (ID NOT IN (SELECT URL from URLs WHERE typeID=4) OR DATEDIFF(dd, datemodified, GETDATE()) = 1)"
rs.Open strSQL, cn
intShopReindex= rs(0)
rs.Close

strSQL = "SELECT count(*) from dsproduct WHERE status<>1 AND ID IN (SELECT URL from URLs WHERE typeID=4)"
rs.Open strSQL, cn
intShopReindex= rs(0) + intShopReindex
rs.Close

strSQL = "SELECT count(*) from Links WHERE live=0"
rs.Open strSQL, cn
intLinks= rs(0) + intLinks
rs.Close

strSQL = "SELECT count(*) from LinkReviews WHERE reviewlive=0"
rs.Open strSQL, cn
intLinks= rs(0) + intLinks
rs.Close

strSQL = "SELECT count(*) from LinkErrors"
rs.Open strSQL, cn
intLinks= rs(0) + intLinks
rs.Close

strSQL = "SELECT count(*) from httperrors WHERE redirectto=''"
rs.Open strSQL, cn
int404 = rs(0)
rs.Close

call getDateShopLastUpdated(cn, rs, strWhen, dteLastUpdated)

strSQL = "EXECUTE sp_spaceused"
rs.Open strSQL, cn
If NOT rs.EOF Then
 dblDBsize = rs("database_size")
Else
 dblDBsize = 0
End If
rs.Close

strSQL = "SELECT count(*) from dsproduct WHERE status=1 AND ID NOT IN (SELECT prodID from dsproductactual)"
rs.open strSQL, cn
If NOT rs.EOF Then
 intProducts = rs(0)
Else
 intProducts = 0
End If
rs.Close

strSQL = "SELECT count(*) from dscategory WHERE ID NOT IN (SELECT catID from dscategoryactual)"
rs.open strSQL, cn
If NOT rs.EOF Then
 intCategories = rs(0)
Else
 intCategories = 0
End If
rs.Close

Session("LoggedIn") = "YES"   'For links manager
Session("admin") = True

blnSilent = False
%> <HTML>

<HEAD>
<META http-equiv="Content-Language" content="en-gb">
<META http-equiv="Content-Type" content="text/html; charset=windows-1252">
<META http-equiv="Refresh" content="120;URL=menu.asp">
<TITLE>Admin Menu</TITLE>
<BASE target="main">
<STYLE>
a{ text-decoration: none; }
</STYLE>
</HEAD>

<BODY topmargin="5" leftmargin="5" link="#000000" vlink="#000000" alink="#AA0000">
<SPAN class="linksin">
<NOBR>
<FORM action="menu.asp" method="GET" target="_self">
<font size="1" face="Verdana">
<INPUT type="hidden" name="compactdb" value="true">

 </font>

<!-- <P align="center"><font face="Verdana" size="1">&nbsp;<B>There are <%=Application("ActiveUsers")%> active users</B></font></P>//-->

<font face="Verdana" size="1">

<%If intForumsReindex > 0 OR intCocktailsReindex > 0 OR intShopReindex > 0 Then%>
</font>
<P><font face="Verdana"><font size="1">&nbsp;</font><B><font size="1" color="#612B83">Just 
Done</font></B><font size="1"><BR>
<A href="/admin/sitesearch/default.asp?action=reindex forums">
<%
If intForumsReindex > 0 Then
Call ReindexSite("reindex forums", "forum post(s)", intForumsReindex, blnSilent)
End If
%>
</a>
<A href="/admin/sitesearch/default.asp?action=reindex cocktails">
<%
If intCocktailsReindex > 0 Then
Call ReindexSite("reindex cocktails", "drink(s)", intCocktailsReindex, blnSilent)
End If
%>
</a>
<A href="/admin/sitesearch/default.asp?action=reindex shop">
<%
If intShopReindex > 0 Then
Call ReindexSite("reindex shop", "product(s)", intShopReindex, blnSilent)
End If
%>
</a>
<%
End If%>
 <%If intNewdrinks>0 OR intNewReviews>0 OR intNewGames>0 OR int404>1 OR intLinks>0 OR DateDiff("d", dteLastUpdated, Now()) <> 0 Then%>
    </font></font>
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
<%if int404>1 Then%> &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="httperrors/default.asp">
 <span style="text-decoration: none">Review 404 errors (<%=int404%> new)</span></A> 
	- <A class="linksin" href="httperrors/default.asp?mode=delete">
 <span style="text-decoration: none">Delete</span></A></B><BR>
 <%End If%>
 <%if intLinks>0 Then%> &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/db/link/links.asp">
 <span style="text-decoration: none">Manage Links (<%=intLinks%>)</span></A></B>&nbsp;<br>
 <%End If%>
 <%If DateDiff("d", dteLastUpdated, Now()) <> 0 Then%> &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/shop/update/updateproducts.asp?Get=true&force=true">
 <span style="text-decoration: none">Update Shop (<%=strWhen%>)</span></A></B>&nbsp;<br>
 <%End If%>
 <%if intProducts>0 AND 1=2 Then%> &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="shop/product.asp">
 <span style="text-decoration: none">Update Products (<%=intProducts%>)</span></A></B><BR>
 <%End If%>
 <%if intCategories >0 AND 1=2 Then%> &nbsp;<font color="#AA0000"></font> 
 </font> <B><A class="linksin" href="shop/category.asp"><font size="1">
 <span style="text-decoration: none">Update Categories (<%=intCategories %>)</span></font></A></B></font></P>
 <font face="Verdana" size="1">
 <%End If%>
 </font>
 <P><font face="Verdana"><font size="1">&nbsp;</font><font color="#612B83" size="1"><B>Managers</B></font><font size="1"><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="sitesearch/default.asp">
 <span style="text-decoration: none">Search Manager</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/admin/db/convert.asp">
 <span style="text-decoration: none">Restore Non-Alcoholic Drinks</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/admin/cocktaileditor">
 <span style="text-decoration: none">Edit Drinks</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="forums/edit_forums.asp">
 <span style="text-decoration: none">Edit Forums</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/db/link/links.asp">
 <span style="text-decoration: none">Edit Links</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="ingredients/modify_ingredients.asp">
 <span style="text-decoration: none">Edit Ingredients</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="ingredients/modify_measures.asp">
 <span style="text-decoration: none">Edit Measures</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/db/member/submitCocktail.asp">
 <span style="text-decoration: none">Add Drink Recipes</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="createHeaderAndFooter.asp">
 <span style="text-decoration: none">Create Header and Footer</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="shop/diagnostic.asp">
 <span style="text-decoration: none">View Shop Xfer Diagnostics</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="shop/searches.asp">
 <span style="text-decoration: none">View Shop Searches</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="shop/randomtest.asp">
 <span style="text-decoration: none">Test Shop Random Products</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="ASPTools/1ClickDBPro/Connect.asp">
 <span style="text-decoration: none">DB Browser</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> </font> <B><A class="linksin" href="http://www.bathandunwind.com/admin/advanced/dir_browser.asp?dir=C%3A%5CInetpub%5Cwwwroot%5Ccocktailuk">
 <font size="1"><span style="text-decoration: none">File Manager</span></font></A></B></font></P>

 <P><font face="Verdana"><font size="1">&nbsp;</font><B><font size="1" color="#612B83">Statistics</font></B><font size="1"><BR>
 &nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <A href="stats/default.asp" class="linksin"><B>
 <span style="text-decoration: none"><%=Capitalise(MonthName(Month(Now), TRUE))%>:</span></B><span style="text-decoration: none"> <%=FormatNumber(intVisits,0)%></span></A><BR>
 &nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <A href="stats/default.asp" class="linksin"><B>
 <span style="text-decoration: none"><%=Capitalise(MonthName(Month(Now), TRUE))%> 
	Pred:</span></B><span style="text-decoration: none"> <%=FormatNumber(CalculateProjectedTarget(intVisits),0)%></span></A><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/admin/sitesearch/searches.asp">
 <span style="text-decoration: none">Referrer Searches</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="menu.asp?compactdb=true" target="_self">
 <span style="text-decoration: none">Current DB Size: <%=dblDBsize%></span></A></B><BR>
 <%If Request("compactdb") = "true" AND Request("newdbsize") = "" Then%> &nbsp;<font color="#AA0000">-</font>-&nbsp;&nbsp; 
 </font></font> <B>
 <font face="Verdana"><font size="1">New DB Size: </font>
 <font size="1" face="Arial"> 
 <INPUT name="newdbsize" value="<%=dblDBSizeNew%>" size="2" style="border: 0px" class="linksin"></font><font size="1">Mb</font></font></B><font face="Verdana"><font size="1">
 </font><font size="1" face="Arial"> <INPUT type="submit" value="Go" style="height: 70%" class="linksin"></font><font size="1"><BR>
 <%ElseIf Request("compactdb") = "true" AND Request("newdbsize") <> "" Then%> &nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <I>
	Database compacted</I></B><BR>
 <%End If%></font></font></P>

<font face="Verdana" size="1">

<%
call GetAffiliateSales(dArrSales, iArrVolume, dArrProfits)
call GetBarmansAffiliateSales(dArrBARAffSales, iArrBARAffVolume, dArrBARAffProfits)
call GetCEAffiliateSales(dArrCEAffSales, iArrCEAffVolume, dArrCEAffProfits)
If blnShowPC Then
  call GetDrinkstuffSales(dArrDSSales, iArrDSVolume, dArrDSProfits)
  call GetBarmansSales(dArrBARSales, iArrBARVolume, dArrBARProfits)
End If
on error goto 0
%> 
 
    </font> 
 
    <P><font face="Verdana"><font size="1">&nbsp;</font><B><font size="1" color="#612B83">Financial</font></B><font size="1"><BR>
 <font color="#AA0000">&nbsp; </font> <B>My Earnings</B><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>T:</B> <%=iArrVolume(0) + iArrCEAffVolume(0) + iArrBARAffVolume(0)%> 
	sales, <%=FormatNumber(dArrProfits(0) + dArrBARAffProfits(0) + dArrCEAffProfits(0), 2)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>Y:</B> <%=iArrVolume(2) + iArrCEAffVolume(2) + iArrBARAffVolume(2)%> 
	sales, <%=FormatNumber(dArrProfits(2) + dArrBARAffProfits(2) + dArrCEAffProfits(2), 2)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B><%=MonthName(Month(Now()), True)%>:</B> <%=iArrVolume(4) + iArrCEAffVolume(4) + iArrBARAffVolume(4)%> 
	sales, <%=FormatNumber(dArrProfits(4) + dArrBARAffProfits(4) + dArrCEAffProfits(4), 2)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>(<%=MonthName(Month(Now()), True)%>):</B> <%=iArrVolume(5) + iArrCEAffVolume(5) + iArrBARAffVolume(5)%> 
	sales, <%=FormatNumber(dArrProfits(5) + dArrBARAffProfits(5) + dArrCEAffProfits(5), 2)%><BR>

 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="http://www.drinkstuff.com/affiliate/default.asp?user=10724&pass=leetracey">
 <span style="text-decoration: none">My C:UK Sales</span></A> - (<a class="linksin" target="_self" href="menu.asp<%If NOT blnShowPC Then%>?percentages=true<%End If%>"><span style="text-decoration: none"><%If NOT blnShowPC Then%>Full<%Else%>Simple<%End If%></span></A>)</B><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>T:</B> <%=iArrVolume(0)%> 
	sales, <%=dArrSales(0)%><%If blnShowPC Then%> (<%=FormatNumber((dArrSales(0)/dArrDSSales(0))*100,0)%>%)<%End If%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>Y:</B> <%=iArrVolume(2)%> 
	sales, <%=dArrSales(2)%><%If blnShowPC Then%> (<%=FormatNumber((dArrSales(2)/dArrDSSales(2))*100,0)%>%)<%End If%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>W:</B> <%=iArrVolume(3)%> 
	sales, <%=dArrSales(3)%><%If blnShowPC Then%> (<%=FormatNumber((dArrSales(3)/dArrDSSales(3))*100,0)%>%)<%End If%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>M:</B> <%=iArrVolume(1)%> 
	sales, <%=dArrSales(1)%><%If blnShowPC Then%> (<%=FormatNumber((dArrSales(1)/dArrDSSales(1))*100,0)%>%)<%End If%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B><%=MonthName(Month(Now()), True)%>:</B> <%=iArrVolume(4)%> 
	sales, <%=dArrSales(4)%><%If blnShowPC Then%> (<%=FormatNumber((dArrSales(4)/dArrDSSales(4))*100,0)%>%)<%End If%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>(<%=MonthName(Month(Now()), True)%>):</B> <%=iArrVolume(5)%> 
	sales, <%=dArrSales(5)%><%If blnShowPC Then%> (<%=FormatNumber((dArrSales(5)/dArrDSSales(5))*100,0)%>%)<%End If%><BR>

    <%If blnShowPC Then%>

    &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="http://www.drinkstuff.com/affiliate/default.asp?user=20724&pass=leetracey">
 <span style="text-decoration: none">My CE Sales</span></A></B><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>T:</B> <%=iArrCEAffVolume(0)%> 
	sales, <%=dArrCEAffSales(0)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>Y:</B> <%=iArrCEAffVolume(2)%> 
	sales, <%=dArrCEAffSales(2)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>W:</B> <%=iArrCEAffVolume(3)%> 
	sales, <%=dArrCEAffSales(3)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>M:</B> <%=iArrCEAffVolume(1)%> 
	sales, <%=dArrCEAffSales(1)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B><%=MonthName(Month(Now()), True)%>:</B> <%=iArrCEAffVolume(4)%> 
	sales, <%=dArrCEAffSales(4)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>(<%=MonthName(Month(Now()), True)%>):</B> <%=iArrCEAffVolume(5)%> 
	sales, <%=dArrCEAffSales(5)%><BR>

    &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="http://www.barmans.co.uk/affiliate/default.asp?user=10724&pass=leetracey">
 <span style="text-decoration: none">My CES Sales</span></A></B><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>T:</B> <%=iArrBARAffVolume(0)%> 
	sales, <%=dArrBARAffSales(0)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>Y:</B> <%=iArrBARAffVolume(2)%> 
	sales, <%=dArrBARAffSales(2)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>W:</B> <%=iArrBARAffVolume(3)%> 
	sales, <%=dArrBARAffSales(3)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>M:</B> <%=iArrBARAffVolume(1)%> 
	sales, <%=dArrBARAffSales(1)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B><%=MonthName(Month(Now()), True)%>:</B> <%=iArrBARAffVolume(4)%> 
	sales, <%=dArrBARAffSales(4)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>(<%=MonthName(Month(Now()), True)%>):</B> <%=iArrBARAffVolume(5)%> 
	sales, <%=dArrBARAffSales(5)%><BR>
 
    &nbsp;<font color="#AA0000"></font> <B><A HREF="http://www.drinkstuff.com/admin/order/" class="linksin">
 <span style="text-decoration: none">Drinkstuff Sales</span></A></B><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>T:</B> <%=iArrDSVolume(0)%> 
	sales, <%=dArrDSSales(0)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>Y:</B> <%=iArrDSVolume(2)%> 
	sales, <%=dArrDSSales(2)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>W:</B> <%=iArrDSVolume(3)%> 
	sales, <%=dArrDSSales(3)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>M:</B> <%=iArrDSVolume(1)%> 
	sales, <%=dArrDSSales(1)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B><%=MonthName(Month(Now()), True)%>:</B> <%=iArrDSVolume(4)%> 
	sales, <%=dArrDSSales(4)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>(<%=MonthName(Month(Now()), True)%>):</B> <%=iArrDSVolume(5)%> 
	sales, <%=dArrDSSales(5)%><BR>

    &nbsp;<font color="#AA0000"></font> <B><A HREF="http://www.barmans.co.uk/admin/order/" class="linksin">
 <span style="text-decoration: none">Barmans Sales</span></A></B><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>T:</B> <%=iArrBARVolume(0)%> 
	sales, <%=dArrBARSales(0)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>Y:</B> <%=iArrBARVolume(2)%> 
	sales, <%=dArrBARSales(2)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>W:</B> <%=iArrBARVolume(3)%> 
	sales, <%=dArrBARSales(3)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>M:</B> <%=iArrBARVolume(1)%> 
	sales, <%=dArrBARSales(1)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B><%=MonthName(Month(Now()), True)%>:</B> <%=iArrBARVolume(4)%> 
	sales, <%=dArrBARSales(4)%><BR>
 	&nbsp;<font color="#AA0000">-</font>&nbsp;&nbsp; <B>(<%=MonthName(Month(Now()), True)%>):</B> <%=iArrBARVolume(5)%> 
	sales, <%=dArrBARSales(5)%><BR>

 <%End If%>

    &nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://www.affiliatewindow.com/">
 <span style="text-decoration: none">AffiliateWindow</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="http://affiliates.williamhill.com/">
 <span style="text-decoration: none">William Hill</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://associates.amazon.co.uk/exec/panama/associates/resources/resources.html/026-0100609-3986040">
 <span style="text-decoration: none">Amazon.co.uk</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://www.tradedoubler.co.uk/pan/login?j_username=me@leetracey.com&j_password=leetracey">
 <span style="text-decoration: none">Trade Doubler</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://adstats.adviva.net/publisher/index.php">
 <span style="text-decoration: none">Adviva</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://www.accelerator-media.com/">
 <span style="text-decoration: none">Accelerator Media</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://www.utarget.co.uk/login.aspx?txtUsername=cocktail&txtPassword=tail71&ddlArea=host">
 <span style="text-decoration: none">UTarget</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <B><A class="linksin " href="http://www.google.com/adsense">
 <span style="text-decoration: none">Google Adsense</span></A></B><BR>
 	&nbsp;<font color="#AA0000"></font> <b>
 <a style="text-decoration: none" href="http://partners.43plc.com">43 Plc</a></b><BR>
 	&nbsp;<font color="#AA0000"></font> </font> <B><A class="linksin " href="accounts/finance.xls">
 <font size="1"><span style="text-decoration: none">Site Finances</span></font></A></B></font></P>
</FORM>
</NOBR>
</HTML>
<%
Response.flush

'Call CreatePrettyURLFiles(cn, rs)

cn.Close
Set cn = Nothing
Set rs = Nothing
%>