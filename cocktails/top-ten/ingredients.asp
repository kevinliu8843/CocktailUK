<%
Option Explicit
strTitle = "Most Popluar Cocktail Ingredients"

Dim cn, dictOKeys, dictOItems, i, j, intNumToDisplay, aryNames
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rs2 = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB
%>
<!--#include virtual="/includes/header.asp" -->

<H2>Top cocktail Ingredients</H2>
<P>These tables should give you some idea of what to purchase first.</P>
<% 
i=0
Do While i<=g_intNumIngredientTypes
	For j=0 to 0

		If i<=g_intNumIngredientTypes Then 
			Response.Write("<H4 align=""center"">" & Capitalise(g_aryIngredientType(i)) & "</H4>")
			Response.Write("<P align=""center"">")
			Call displayTopTen(g_aryIngredientTypeID(i))
		Else
			Response.Write("&nbsp;")
		End If

		i=i+1
	Next 
Loop
%>
</TABLE>
</td>
  </tr>
</table><%
cn.Close
Set cn = Nothing
Set rs = Nothing
Set rs2 = Nothing

' Routines go here --------------------------------------------------------

Sub displayTopTen(strType)
	Dim i

	Call GetTopTenIngredients(strType, aryNames, intNumToDisplay)
	intNumToDisplay = intNumToDisplay - 1
%>
<TABLE border="1" cellpadding="2" cellspacing="0" bordercolor="#FFFFFF" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" bgcolor="#612b83" style="border-collapse: collapse">
  <TR>
    <Th width="10%">Position</Th>
    <Th width="80%">Name</Th>
    <Th width="10%">Used</Th>
  </TR>
	<%For i=0 to intNumToDisplay%>
  <TR>
    <TD><FONT color="#FF0000" size="3"><%=i+1%></FONT></TD>
    <TD><%=Capitalise(aryNames(0, i))%></TD>
    <TD><%=aryNames(1, i)%></TD>
  </TR>
	<%Next%>
</TABLE>
<%
End Sub

Sub GetTopTenIngredients(strType, aryNames, intNumToDisplay)
	strSQL = "SELECT DISTINCT TOP 10 Ingredients.name, COUNT(CocktailIng.ingredientID) AS CountIng"
	strSQL = strSQL & " FROM         CocktailIng INNER JOIN"
	strSQL = strSQL & " 	ingredients ON CocktailIng.ingredientID = ingredients.ID INNER JOIN"
	strSQL = strSQL & "     cocktail ON CocktailIng.cocktailID = cocktail.ID"
	strSQL = strSQL & " 	WHERE     (ingredients.type = "&strtype&") AND (cocktail.status=1)"
	strSQL = strSQL & " 	GROUP BY Ingredients.name"
	strSQL = strSQL & " 	ORDER BY CountIng DESC"
	rs.open strSQL, cn, 0, 3
	aryNames=  rs.GetRows()
	intNumToDisplay = 10
	rs.close
End Sub
%><!--#include virtual="/includes/footer.asp" -->