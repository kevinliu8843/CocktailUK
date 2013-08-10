<%
Option Explicit
strTitle="What Can I Make?"
Dim strIDS, cn

If NOT Session("logged") Then
	Response.Redirect ("/account/login.asp")
End If
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<%
set cn	= Server.CreateObject("ADODB.Connection")
Set rs	= Server.CreateObject("ADODB.RecordSet")
cn.Open strDBMod
strSQL = "EXECUTE CUK_RECIPESUSERCANMAKE @m="&Session("ID")
call writeCocktailList(strSQL, rs, cn, "What Can I Make?", "")
%>
<!--#include virtual="/includes/footer.asp" -->