<%
Option Explicit

Dim cn, aryRows, i, intNumReviews, aryIng, intNumIng

strTitle = "Accept/Decline Reviews"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<H2>Reviews awaiting verification</H2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
<FORM method="POST" action="default.asp">
<%If Request("editing") <> "" Then%>
  <INPUT type="hidden" name="ingID" value="<%=Request("editing")%>">
<%End if%>
<%If Request("edit") <> "" Then%>
  <INPUT type="hidden" name="ID" value="<%=Request("edit")%>">
<%End if%>

<%
set cn = Server.CreateObject("ADODB.Connection")
cn.Open strDBMod

If Request("ID") <> "" Then
	cn.execute("UPDATE cocktailreview Set review='"&strIntoDB(Request("review"))&"' WHERE ID="&Request("ID"))
End if

If Request("ingID") <> "" Then
	cn.execute("UPDATE drink_desc Set description='"&strIntoDB(Request("review"))&"' WHERE ID="&Request("ingID"))
End if

If Request.QueryString("delete") <> "" Then
	Set rs = cn.Execute("UPDATE cocktailreview Set status=2 WHERE ID=" & Int(Request.QueryString("delete")))
End If

If Request.QueryString("deleteing") <> "" Then
	Set rs = cn.Execute("UPDATE drink_desc Set status=2 WHERE ID=" & Int(Request.QueryString("deleteing")))
End If

If Request.QueryString("accept") <> "" Then
	cn.Execute("UPDATE cocktailreview SET status=1 WHERE status=0")
	cn.Execute("UPDATE drink_desc SET status=1 WHERE status=0")
End If

Set rs = cn.Execute ("SELECT name, ID, review, cocktailID from Cocktailreview WHERE Status=0")
If NOT rs.EOF Then
	aryRows = rs.GetRows()
	intNumReviews = UBound(aryRows,2)
Else
	ReDim aryRows(1,1)
	intNumReviews = -1
End If
rs.close

rs.open "SELECT name, ID, description, drink_id from drink_desc WHERE status=0", cn, 0, 3
If NOT rs.EOF Then
	aryIng = rs.getRows()
	intNumIng = UBound(aryIng,2)
Else
	ReDim aryIng(1,1)
	intNumIng = -1
End If
rs.close

If intNumReviews > 0 Then
	Response.write "  <H4>Drink reviews</H4>"
End If
For i=0 To intNumReviews
	rs.open "SELECT name from cocktail WHERE ID="&aryRows(3,i), cn, 0, 3
	Response.Write "Recipe: " & rs("name") 
	If CStr(Request("edit") & "") = CStr(aryRows(1,i)) then
		Response.write ", Person: " & Server.HTMLEncode(aryRows(0,i)) & " - (<A href=default.asp?delete="& aryRows(1,i) &">Delete</a>)"  & "<BR><TEXTAREA name=review cols=50 rows=5>"&Server.HTMLEncode(Replace(ReplaceStuff(aryRows(2, i))), "<BR>", VbCrLf) & "</TEXTAREA>" & VbCrLf
		Response.write "<BR><CENTER><INPUT type=""submit"" value=""Update""></CENTER><HR>"
	Else
		Response.write ", Person: " & aryRows(0,i) & " - (<A href=default.asp?delete="& aryRows(1,i) &">Delete</a> | <A href=default.asp?edit="& aryRows(1,i) &">Edit</a>)"  & "<BR><IMG border=""0"" src=""/images/inset_quotebegin.gif"" width=14 height=10> "&Server.HTMLEncode(aryRows(2,i)) & " <IMG border=""0"" src=""/images/inset_quoteend.gif"" width=14 height=10><HR>" & VbCrLf
	End If
	
	rs.close
Next
If intNumIng > 0 Then
	Response.write "  <H4>Ingredient reviews</H4>"
End If 
For i=0 To intNumIng 
	rs.open "SELECT name from ingredients WHERE ID="&aryIng(3,i), cn, 0, 3
	Response.Write "Ingredient: " & rs("name")
	If CStr(Request("editing") & "") = CStr(aryIng(1,i)) then
		Response.write ", Person: " & Server.HTMLEncode(aryIng(0,i)) & " - (<A href=default.asp?deleteing="& aryIng(1,i) &">Delete</a>)"  & "<BR><TEXTAREA name=review cols=50 rows=5>"&Server.HTMLEncode(Replace(ReplaceStuff(aryIng(2,i)), "<BR>", VbCrLf)) & "</TEXTAREA>" & VbCrLf
		Response.write "<BR><CENTER><INPUT type=""submit"" value=""Update""></CENTER><HR>"
	Else
		Response.write ", Person: " & aryIng(0,i) & " - (<A href=default.asp?deleteing="& aryIng(1,i) &">Delete</a> | <A href=default.asp?editing="& aryIng(1,i) &">Edit</a>)"  & "<BR><IMG border=""0"" src=""/images/inset_quotebegin.gif"" width=14 height=10> "&Server.HTMLEncode(aryIng(2,i)) & " <IMG border=""0"" src=""/images/inset_quoteend.gif"" width=14 height=10><HR>" & VbCrLf
	End If
	
	rs.close
Next
cn.close
Set cn =nothing
Set rs = nothing
%> <CENTER>
<p><INPUT class="button" type="button" value="Accept reviews &gt; &gt;" onClick="location.href='default.asp?accept=true'">
</p>
</CENTER>
</FORM>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->