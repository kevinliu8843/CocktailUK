<%
Option Explicit
Dim cn, strRating, strID
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
strRating = Request("R1")
strID = Request("ID")
If NOT IsNumeric(strID) OR NOT IsNumeric(strRating) Then
	Response.redirect("/")
End If

If strRating = "" Then
	If Request.ServerVariables("HTTP_REFERER") <> "" Then
		Response.Redirect(Request.ServerVariables("HTTP_REFERER") & "&rate=false")
	Else
		Response.redirect("/default.asp")
	End If
End If
set cn = Server.CreateObject("ADODB.Connection")
cn.Open strDBMod

IF strID <> "" AND IsNumeric(strID) AND IsNumeric(strRating) Then
	If Request("game") = "true" Then
		Set rs2 = cn.Execute("SELECT * FROM drinkinggame WHERE Status=1 And ID=" & strIntoDB(strID))
		addGameRating rs2("peoplerated"), rs2("rating"), Int( strRating ), strID, cn
	Else
		Set rs2 = cn.Execute("SELECT * FROM cocktail WHERE Status=1 And ID=" & strIntoDB(strID))
		addRating rs2("users"), rs2("rate"), Int( strRating ), strID, cn
	End If
Else
	cn.close
	Set cn = nothing
	response.Redirect("/default.asp")
End If
%>
