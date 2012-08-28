<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%strTitle="View shop searches"%>
<!--#include virtual="/includes/header.asp" -->
 <h2>View shop product searches</h2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">Here is a list of user typed searches in order of how many times searched...
  <FORM method="POST" action="searches.asp?delete=true">
   <P align="center"><INPUT type="submit" value="Click here to delete searches from database &gt; &gt;" name="B3" class="button"></P>
 </FORM>
<BLOCKQUOTE>
<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open strDB

If Request("delete") = "true" Then
	cn.execute ("DELETE from dssearches WHERE site=0")
End If

strSQL = "SELECT search, count(search) as count from dssearches WHERE site=0 GROUP BY search order by count(search) DESC"

rs.open strSQL, cn, 0, 3
While NOT rs.EOF
	Response.write "<A HREF=""/shop/products/search.asp?search="&strOutDB(Replace(rs("search"), "%", ""))&""" target=""_blank"">" & strOutDB(Replace(rs("search"), "%", " ")) & "</A> ("&rs("count")&")<BR>"
	rs.movenext
WEND
rs.Close

cn.Close
Set cn = Nothing
Set rs = Nothing
%>
</BLOCKQUOTE>
    <p>&nbsp;</td>
    </tr>
  </table>
<!--#include virtual="/includes/footer.asp" -->