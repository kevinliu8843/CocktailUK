<%@ Language=VBScript %>
<!--#include virtual="/includes/variables.asp" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
</HEAD>
<BODY><TEXTAREA style="HEIGHT: 565px; WIDTH: 595px" rows="1" cols="20">
<%

Server.ScriptTimeout = 2
Set cn = Server.CreateObject("ADODB.Connection")
Set dict = Server.CreateObject("scripting.dictionary")
Set dict2 = Server.CreateObject("scripting.dictionary")
Set dict3 = Server.CreateObject("scripting.dictionary")
cn.Open strDB
Set rs = cn.Execute ("SELECT DISTINCT ID from cocktail")

Do While NOT rs.EOF
	dict.Add CStr(rs("ID")), CStr(rs("ID"))
	rs.MoveNext
Loop

dict3.Add CStr(33), CStr(33)
dict3.Add CStr(18), CStr(18)
dict3.Add CStr(17), CStr(17)
dict3.Add CStr(20), CStr(20)

Set rs = Nothing
cn.Close
Set cn = Nothing

upperbound = dict.Count
lowerbound = 1
For i=1 to 365
	Randomize
	rndID = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
	Response.Write VbTab & VbTab & "case " & i & " getCOWID=" & rndID & " ' " & DateAdd("d", i-1, CDate("1/1/"&Year(Now()))) & VbCrLF
Next

%>
</TEXTAREA>
</BODY>
</HTML>