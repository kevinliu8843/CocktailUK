<%
Option Explicit
Dim cn, strIDS, upperbound, lowerbound, arrMatch
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDBMod
Set rs = cn.Execute("EXECUTE CUK_RECIPESUSERCANMAKE @m="&Session("ID"))
If NOT rs.EOF Then
	arrMatch = rs.Getrows
	Set rs=  nothing
	cn.close
	Set cn = nothing
	If IsArray(arrMatch) Then
		upperbound = UBound(arrMatch,2)
		lowerbound = 0
		Response.Redirect("/db/viewCocktail.asp?ID=" & arrMatch( 0, randomise( upperbound, lowerbound ) ))
	Else
		Response.Redirect("/db/viewCocktail.asp?ID=" & randomise(0, 9000))
	End If
Else
	Set rs=  nothing
	cn.close
	Set cn = nothing
	response.Redirect("/db/viewCocktail.asp?ID=" & randomise(0, 9000))
End If

Function randomise( upperbound, lowerbound )
	'Returns an ID to be checked
	Randomize
	randomise = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function
%>
