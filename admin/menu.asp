<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
'On Error Resume Next

Session("admin") = True

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

cn.Open strDB

rs.open "SELECT COUNT(*) FROM usr",cn
intUsers = rs(0)
rs.close

rs.open "SELECT COUNT(*) FROM cocktail WHERE status=1",cn
intRecipes = rs(0)
rs.close

cn.Close
Set cn = Nothing
Set rs = Nothing

intNewdrinks = 0
%>
<HTML>

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

<div style="margin: auto"><%=intUsers%> Users | <%=intRecipes%> Recipes</div>
<SPAN class="linksin">
 <%If intNewdrinks>0 Then%>
    <P><font face="Verdana"><font size="1">&nbsp;</font><B><font size="1" color="#612B83">To 
	Do List</font></B><font size="1"><BR>
 <%End If%>
 <%if intNewdrinks>0 Then%> &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="submitCocktails/default.asp">
	 <span style="text-decoration: none">Add Users' Drinks (<%=intNewdrinks%> new)</span></A></B><BR>
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
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/account/submitCocktail.asp">
 <span style="text-decoration: none">Add Drink Recipes</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="createHeaderAndFooter.asp">
 <span style="text-decoration: none">Create Header and Footer</span></A></B><BR>
 &nbsp;<font color="#AA0000"></font> <B><A class="linksin" href="/admin/sitemap.asp">
 <span style="text-decoration: none">Rebuild recipes/categories</span></A></B><BR>
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
