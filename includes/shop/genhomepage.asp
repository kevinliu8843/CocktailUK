<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/shop.asp" -->
<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open strDB
Call GenerateHomePage(cn, rs)
cn.close
Set cn = Nothing
Set rs = Nothing
%>