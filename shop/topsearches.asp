<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%strTitle="Top 100 shop searches"%>
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" -->
 <h2>Cocktail : UK top 100 shop searches</h2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">Here is a list what the most popular searches are on the 
	Cocktail : UK shop...
<OL>
<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open strDB

strSQL = "SELECT Top 100 search from dssearches GROUP BY search order by count(search) DESC"
rs.open strSQL, cn, 0, 3
While NOT rs.EOF
	Response.write "<LI><A HREF=""/shop/products/search.asp?search="&Replace(strOutDB(rs("search")), "%", "")&""">" & strOutDB(rs("search")) & "</A></LI>"
	rs.movenext
WEND
rs.Close

cn.Close
Set cn = Nothing
Set rs = Nothing
%>
</OL>
</td>
    </tr>
  </table>
<!--#include virtual="/includes/footer.asp" -->