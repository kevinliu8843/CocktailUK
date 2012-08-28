<%strtitle="Set up non-alcoholic drinks"%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
Set cn	= Server.CreateObject("ADODB.Connection")
cn.open strDBMod
cn.execute( "EXECUTE CUK_RESETNONALCOHOLICRECIPES")
Set cn	= Nothing
%>
<P>Non-Alcoholic recipes reset...<!--#include virtual="/includes/footer.asp" -->