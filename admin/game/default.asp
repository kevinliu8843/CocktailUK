<%
Option Explicit
Dim cn, aryRows, i, intNumReviews, intNumGames
strTitle = "Accept/Decline Games"
On Error Resume Next
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<H2>Games awaiting verification</h2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
<%
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open strDBMod

If Request.QueryString("delete") <> "" Then
	Set rs = cn.Execute("UPDATE drinkinggame Set status=2 WHERE ID=" & Int(Request.QueryString("delete")))
End If

If Request.QueryString("accept") <> "" Then
	cn.Execute("UPDATE drinkinggame SET status=1 WHERE status=0")
End If

Set rs = cn.Execute ("SELECT * from drinkinggame WHERE Status=0")
While NOT rs.EOF
	Response.Write "Game title: " & strOutDB(rs("title")) & ", Person: " & strOutDB(rs("submitter")) & ", Type: "
	Response.Write aryGames(rs("type"))
	Response.Write " - (<A href=default.asp?delete="& rs("ID") &">Delete Game</a>)"  & "<BR><IMG border=""0"" src=""/images/inset_quotebegin.gif""> "& strOutDB(rs("directions")) & " <IMG border=""0"" src=""/images/inset_quoteend.gif""><HR>" & VbCrLf
	rs.movenext
Wend
cn.close
Set cn = nothing
Set rs = nothing
%>
<CENTER>
<p><INPUT  class="button" type=button value="Accept games &gt; &gt;" onClick="location.href='default.asp?accept=true'"></p>
    </CENTER></td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->