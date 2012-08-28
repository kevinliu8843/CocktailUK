<%
Option Explicit
strTitle = "Top Ten Ingredients"

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

<H2>Our most popular ingredients</H2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
<P>These tables should give you some idea of what to purchase first.</P>
<TABLE border="0" cellpadding="0" cellspacing="10" width="100%">
<% 
i=0
Do While i<=g_intNumIngredientTypes
	Response.Write("<TR>")
	For j=0 to 0
	    Response.Write("<TD width=""100%"" valign=""top"" align=""left"">")

		If i<=g_intNumIngredientTypes Then 
			Response.Write("<H4 align=""center"">" & Capitalise(g_aryIngredientType(i)) & "</H4>")
			Response.Write("<P align=""center"">")
			Call displayTopTen(g_aryIngredientTypeID(i))
		Else
			Response.Write("&nbsp;")
		End If

		Response.Write("</TD>")
		i=i+1
	Next 
	Response.Write("</TR>")
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
    <TD bgcolor="#612b83"><FONT color="white"><B>Position</B></FONT></TD>
    <TD width="100%" bgcolor="#612b83"><FONT color="white"><B>Name</B></FONT></TD>
    <TD width="50%" bgcolor="#612b83"><FONT color="white"><B>Used</B></FONT></TD>
  </TR>
	<%For i=0 to intNumToDisplay%>
  <TR>
    <TD bgcolor="#CCCCCC"><FONT color="#FF0000" size="3"><%=i+1%></FONT>&nbsp;</TD>
    <TD width="100%" bgcolor="#CCCCCC"><%=Capitalise(aryNames(0, i))%>&nbsp;</TD>
    <TD width="50%" bgcolor="#CCCCCC"><%=aryNames(1, i)%>&nbsp;</TD>
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