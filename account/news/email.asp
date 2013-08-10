<%
Option Explicit

Dim cn
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open strDB
strSQL = "SELECT email from usr where news=true"
Set rs = cn.Execute( strSQL )

Do While NOT rs.EOF
	Response.Write(rs("email") &"; " & VbCrLf)
	rs.MoveNext
Loop

Set rs = Nothing
cn.Close
Set cn = Nothing
%>