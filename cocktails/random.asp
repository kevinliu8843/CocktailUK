<%
Option Explicit

Dim cn, upperbound, lowerbound, ID

If Session("logged") Then
	Response.Redirect("/account/random.asp")
End If

%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%

'Random ID generator
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB

strSQL = "SELECT count(*) FROM cocktail WHERE Status=1"
rs.Open strSQL, cn, 3, 3

upperbound = rs(0)
lowerbound = 1

ID = randomise ( upperbound, lowerbound )

rs.Close
Set rs = Nothing
cn.Close
Set cn = Nothing

response.Redirect("/cocktails/recipe.asp?ID=" & ID)

Function randomise( upperbound, lowerbound )
	'Returns an ID to be checked
	Randomize 
	randomise = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function
%>