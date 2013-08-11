<%
Option Explicit
%>

<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" -->
<%
'Finds all cocktails containing a particular ingredient

'Get URL string INFO
Dim strHrefType, strIngredient, blnIngredient
Dim cn, strSQLName, strDrink, strSQLStore

strHrefType = ""
strIngredient = Request.QueryString("ingredient")
If NOT Isnumeric(strIngredient) Then 
	strIngredient=  "1"
End If
blnIngredient = False

'Determine type of query and setup SQL query string
If strIngredient <> ""  Then
	blnIngredient = True
	strSQLStore = "SELECT cocktail.ID, cocktail.type, cocktail.name, ingredients.name As IngName FROM CocktailIng INNER JOIN cocktail ON CocktailIng.cocktailID = cocktail.ID INNER JOIN ingredients ON CocktailIng.ingredientID = ingredients.ID WHERE Cocktail.status=1 AND (CocktailIng.ingredientID = "&strIngredient&")"
	strHrefType = "&ingredient="&strIngredient
Else
	If Request.ServerVariables("HTTP_REFERER") <> "" Then
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Else
		Response.Redirect("/")
	End If
End If

'-------------------------------------------------------------

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB

rs.Open strSQLStore, cn, 0, 3
if rs.eof then
	rs.close
	cn.Close 
	Set cn = Nothing
	Set rs = Nothing
	If Request.ServerVariables("HTTP_REFERER") <> "" Then
		Response.Redirect(Request.ServerVariables("HTTP_REFERER") & "?error=" & Server.HtmlEncode("Sorry, no cocktails contain " & strDrink ))
	Else
		Response.Redirect("/")
	End If
Else
	strDrink = rs("IngName")
	strTitle = "cocktails containing " & strDrink
	%>
	<!--#include virtual="/includes/header.asp" -->
	<%
	rs.Close
	Call writeCocktailList(strSQLStore, rs, cn, strTitle, strHrefType)
End if

cn.Close 
Set cn = Nothing
Set rs = Nothing
%><!--#include virtual="/includes/footer.asp" -->