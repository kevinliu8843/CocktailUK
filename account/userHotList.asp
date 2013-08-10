<%
Option Explicit
strTitle="My Favourites"
Dim cn, strUserList, blnListOk, i, blnDuplicated, strPrint, extraQueryString

If NOT Session("logged") Then
	If Request.QueryString("add") <> "" Then
		extraQueryString = "?add=" & Request.QueryString("add")
	End If
	response.Redirect("/account/loginOut.asp" & extraQueryString)
End If

' OK, we are logged in, let's proceed...
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" -->
<%
set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDBMod


'Addition code----------------------------------------------------------------------------------------------------

If Request.QueryString("add") <> "" AND IsNumeric(Request.QueryString("add")) Then
	'add a cocktail to the list
	strSQL = "INSERT into usrfav (memID, cocktailID) VALUES (" & Session("ID")&", "&Request.QueryString("add") & ")"
	strPrint = "<p><FONT color=red><i>Cocktail added to your favourites</i></font>"
	cn.execute(strSQL)
End If

'Removal code----------------------------------------------------------------------------------------------------

If Request.QueryString("remove") <> "" AND IsNumeric(Request.QueryString("remove")) Then
	'add a cocktail to the list
	strSQL = "DELETE from usrfav WHERE memID="&Session("ID")&" AND cocktailID="&Request.QueryString("remove") 
	cn.execute(strSQL)
	strPrint = "<p><FONT color=red><i>Cocktail removed from your favourites</i></font>"
End If
%>

<!--#include virtual="/includes/header.asp" -->
<%
strSQL = "SELECT count(*) FROM cocktail WHERE Status=1 And ID IN (SELECT cocktailID from usrfav where memID="&Session("ID")&")"
rs.open strSQL, cn, 0, 3
blnListOk = NOT rs.EOF
rs.close
IF blnListOk Then
	strSQL = "SELECT DISTINCT name, ID, type FROM cocktail WHERE Status=1 And ID IN (SELECT DISTINCT cocktailID from usrfav where memID="&Session("ID")&") order by name"
	strTitle = Session("firstName") &"'s favourite cocktails"
	Call writeCocktailList(strSQL, rs, cn, strTitle, "")
	If strPrint <> "" then
		Response.write strPrint
	End If
Else
%>
<p>
<p>You don't appear to have any cocktails saved in your favourites.
<P>
To start a list of favourites, find some cocktails and click "save this cocktail in your favourites" on the cocktail page 
<%
End If

cn.Close
Set cn = Nothing
Set rs = Nothing
%><!--#include virtual="/includes/footer.asp" -->