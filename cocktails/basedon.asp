<%
Option Explicit

Dim cn, basedID, strHrefType, strSQLcocktail, basedOn, based, iPageCurrent, iPageSize, iPageCount
Dim FSO, name, FileExists, strType, strAddType, strOrderBy
%>
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" -->
 
<%
set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB

'--------------------------------------------------------
basedID = Request("basedID")

'Error trap
If basedID = "" OR NOT IsNumeric(basedID) Then basedID = 1

strHrefType = "&basedID="&strIntoDB(basedID)

strSQLcocktail = "SELECT name, ID, type FROM cocktail WHERE Status=1 And based=" & basedID

If CStr( basedID ) = "5" Then	'include dark rum based drinks too
	strSQL = "SELECT name, ID FROM ingredients WHERE ID=" & basedID & " OR ID=6 OR ID=7"
	strSQLcocktail = "SELECT name, ID, type FROM cocktail WHERE Status=1 And based IN (" & basedID & ",6,7)"
	based = "Rum"
Else
	strSQL = "SELECT name, ID FROM ingredients WHERE ID="&basedID
End If

If CStr( basedID ) = "8" Then	'include gold tequila based drinks too
	strSQL = "SELECT name, ID FROM ingredients WHERE ID=" & basedID & " OR ID=9"
	strSQLcocktail = "SELECT name, ID, type FROM cocktail WHERE Status=1 And based IN (" & basedID & ",9)"
	based = "Tequila"
Else
	strSQL = "SELECT name, ID FROM ingredients WHERE ID="&basedID
End If

If CStr( basedID ) = "4" Then	'include other whiskies too
	strSQL = "SELECT name, ID FROM ingredients WHERE ID IN (" & basedID & ",10,15)"
	strSQLcocktail = "SELECT name, ID, type FROM cocktail WHERE Status=1 And based IN (" & basedID & ",10,15)"
	based = "Whisky"
Else
	strSQL = "SELECT name, ID FROM ingredients WHERE ID="&basedID
End If

If Request("orderby") <> "" Then
	Session("orderby") = Request("orderby")
End If
If Session("orderby") <> "" Then
	If Session("orderby") = "name" Then
		strOrderBy = " ORDER BY name ASC"
	ElseIf Session("orderby") = "rate" then
		strOrderBy = " ORDER BY rate DESC, accessed DESC"
	Else
		strOrderBy = " ORDER BY " & Session("orderby") & " DESC, name ASC"
	End If
Else
	strOrderBy = " ORDER BY accessed DESC"
	Session("orderby") = "accessed"
End If

If Request("naughty") = "" Then
	strAddType = strAddType & " AND type<>(type | 8)"
Else
	strHrefType = strHrefType & "&naughty=ON"
End If

If Request("userrecipes") = "" Then
	strAddType = strAddType & " AND Len(usr)<=0"
Else
	strHrefType = strHrefType & "&userrecipes=ON"
End If

rs.Open strSQL, cn, 0, 3
If Not rs.EOF Then
	basedOn = rs("name")
End If
rs.Close
If NOT CStr( basedID ) = "5" Then
	based = capitalise( replaceStuffBack( basedOn ) )
End If

'--------------------------------------------------------
%>
<!--#include virtual="/includes/header.asp" -->
<%
Select Case basedID
	Case "2"
	%>
		<h1>Brandy Based Cocktails - Eau-de-Vie-de-Vin</h1>
		<P style="padding-bottom: 20px;">Brandy is a warm spirit
		distilled from wine and is produced in many forms from around the world.
		The best known are the Cognacs and Armagnacs of France. However, good
		brandy is also produced by many other countries around the world. Top-quality brandies are
		produced in copper pot stills. The tradition of maturing brandy goes back
		to the 15<SUP>th </SUP>century, when an alchemist allegedly took his
		precious barrel of aqua vitae and buried it in his yard to protect it
		against the guards and soldiers. The barrel was found years later but half
		of the liquor had evaporated leaving a rich smooth nectar. Brandy is an extremely
		versatile drink that mixes well with a wide range of other flavours, as the
		following recipes demonstrates.</p>
	<%
	Case "3"
	%>
		<h1>Gin Based Cocktails - Mother's Ruin</h1>
		<P style="padding-bottom: 20px;">Gin is a very common used
		base for cocktails, including the famous martini. The original gin was
		concocted illegally in bathtubs and bore little resemblance to the gin
		we know today. It was probably this poor flavour that gave cocktails their
		popularity. Gin is a &quot;neutral,
		rectified spirit distilled from any grain, potato or beet and flavoured
		with juniper&quot;. Clearly with this definition there are alot of
		variations of &quot;gin&quot; around. And each manufacturer keeps their
		recipe closely guarded.</p>
	<%
	Case "5"
	%>
		<h1>Rum Based Cocktails - Kill-Devil</h1>
		<P style="padding-bottom: 20px;">Rum is a rich and fragrant
		spirit, distilled from molasses in a pot still or a patent still. The drink
		is associated with the tropics because most sugar cane is grown there
		(this is where molasses come from). Rum is clear when is comes from the
		still (white rum), but is also matured in oak casks to give it a caramel colour
		(dark rum). Drinkers
		have discovered that the rich taste combines perfectly with fruit
		juice to give a tropical flavour to cocktails. There
		are hundreds of rum-based cocktails available to mixologists, however
		here are a select few.</p>
	<%
	Case "1"
	%>
		<h1>Vodka Based Cocktails - Zhiznennia Voda</h1>
		<P style="padding-bottom: 20px;">The word &quot;vodka&quot;
		comes from the Russian &quot;Zhiznennia voda&quot;, which means &quot;water
		of life&quot;. &quot;Vodka&quot; means &quot;little water&quot;.Vodka
		can be distilled from any number of sources including potatoes and sugar
		cane. The spirit is then filtered through charcoal to filter out
		impurities, and the resultant liquor is then the perfect base for a
		cocktail. It adds the kick of alcohol but without the taste. This is one
		of the great things about vodka: it is almost odourless. So it cannot be
		detected on your breath. Vodka is the ideal drink for beginners to use in experiments.</p>
	<%
	Case "4"
	%>
		<h2>Whisky Based Cocktails - The Hard Stuff</h2>
		<P style="padding-bottom: 20px;">Whisky (or whiskey -
		depending on its origin) is made in many parts of the world. Whisky
		(without the &quot;e&quot;) is produced in Scotland. This is usually
		referred to as Scotch. All other whiskies are spelled with an
		&quot;e&quot; to distinguish them from the real thing. Essentially
		whisky is distilled from fermented mash made of malted grain. In Scotland
		it derives much of its rich flavour from peat. Some
		of the best known are single malts, which are produced from a particular
		distillery. While most commercial ones are blended, made from products of several regions.</p>
	<%
	Case Else
		strTitle = "List of " & based & " based cocktails"
End Select

Call writeCocktailList( strSQLcocktail & strAddType & strOrderBy, rs, cn, strTitle, strHrefType )
cn.Close
Set cn = Nothing
Set rs = Nothing
%>
<FORM name="order" action="basedon.asp" method="GET">
  <div align="center">
  <table border="0" cellpadding="2" cellspacing="0" id="table1">
	<tr>
		<td valign="top"><B>Order By :</B></td>
		<td valign="top"><B>&nbsp;<INPUT type="radio" name="orderby" value="accessed" onclick="order.submit()" <%If Session("orderby")="accessed" then%>checked<%End If%> id="fp1" checked></B><LABEL for="fp1">Times Viewed</LABEL><B> 
  <INPUT type="radio" value="name" name="orderby" onclick="order.submit()"  <%If Session("orderby")="name" then%>checked<%End If%> id="fp2"></B><LABEL for="fp2">Name</LABEL><B> 
  <INPUT type="radio" name="orderby" value="rate" onclick="order.submit()" <%If Session("orderby")="rate" then%>checked<%End If%> id="fp3"></B><label for="fp3">Rating</label></td>
	</tr>
	<tr>
		<td valign="top"><b>Include:</b></td>
		<td valign="top">
  <P align="left">
		<input type="checkbox" name="userrecipes" value="ON" id="fp5" onclick="order.submit()" <%If Request("userrecipes")="ON" Then%> CHECKED<%End If%>><label for="fp5">User submitted recipes</label></P>
  		</td>
	</tr>
	</table>
  </div>
  <INPUT type="hidden" name="basedID" value="<%=basedID%>">
</FORM>
<!--#include virtual="/includes/footer.asp" -->