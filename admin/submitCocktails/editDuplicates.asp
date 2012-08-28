<%
Option Explicit

CONST NO_MATCH			= 0
CONST INGREDIENT_MATCH	= 1
CONST BOTH_MATCH		= 2

Dim cn, intID, strSimilar, arySimilar, i, strAction, strDeleteIDs
%>
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/variables.asp" -->

<SCRIPT language="javascript">
window.focus();
</SCRIPT>

<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDBMod
intID			= Request("ID")
strSimilar		= Request("similar")
strAction		= Request("action")
strDeleteIDs	= Request("delete")

If strAction = "delete" And strDeleteIDs <> "" Then
	strSQL = "UPDATE Cocktail SET Status=2, ReIndex=1 WHERE ID IN (" & strDeleteIDs & ")"
	cn.Execute(strSQL)
End If

Response.Write("<form action=""editduplicates.asp"" method=""post"">")
Response.Write("<input type=""hidden"" name=""ID"" value=""" & intID & """>")
Response.Write("<input type=""hidden"" name=""similar"" value=""" & Server.HTMLEncode(strSimilar) & """>")
Response.Write("<input type=""submit"" name=""action"" value=""delete"">")
Response.Write("<table>")


Response.Write("<tr><td width=""50%"">")
Call DisplayCocktail(intID)
Response.Write("</td><td valign=""top"" width=""50%"">")
Response.Write("<input type=""checkbox"" name=""delete"" value=""" & intID & """>")
Response.Write("</td></tr>")



If Instr(strSimilar, "|") Then
	arySimilar = Split(strSimilar, "|")
Else
	ReDim arySimilar(0)
	arySimilar(0) = strSimilar
End If


For i=0 To UBound(arySimilar)
	If Int(Split(arySimilar(i), ",")(1)) = BOTH_MATCH Then
		Response.Write("<tr><td width=""50%"" bgcolor=""red"">")
	Else
		Response.Write("<tr><td width=""50%"" bgcolor=""green"">")
	End If

	Call DisplayCocktail(Split(arySimilar(i), ",")(0))
	Response.Write("</td>")

	Response.Write("<td valign=""top"" width=""50%"">")
	Response.Write("<input type=""checkbox"" name=""delete"" value=""" & Split(arySimilar(i), ",")(0) & """>")
	Response.Write("</td></tr>")
Next


Response.Write("<table></form>")



Sub DisplayCocktail(intID)
	Dim strIngredients, strRecipe, strName, strDescription
	Dim MyArray, MyArray2, x, y, i

	strSQL = "SELECT * FROM Cocktail WHERE Status<>2 And ID=" & intID

	rs.Open strSQL, cn, 0, 3
	If Not rs.EOF Then
		strIngredients	= rs("ingredients")
		strName			= rs("Name")
		strDescription	= rs("Description")
	End If
	rs.Close

	If strIngredients <> "" Then
		' Get Ingredients --------------------------------------------------------
		strRecipe = ""
		If strIngredients <> "" Then
			MyArray = Split(CStr( strIngredients ), ",", -1, 1)
			For Each x in MyArray
				MyArray2 = Split(Cstr( x ), "-", -1, 1)
				i = 0
				For Each y in MyArray2
					i = i + 1
					If (Int(i) Mod 2) = 1 Then 'array part is the amount
						strSQL = "SELECT name, ID FROM measure WHERE ID=" & y
						rs.Open strSQL, cn, 0, 3
						If NOT rs(0) = "no measure" Then
							strRecipe = strRecipe & rs(0) & " "
						Else
							strRecipe = strRecipe & " "
						End If
						rs.Close
					Else				'array part is the type
						strSQL = "SELECT name, ID, type FROM ingredients WHERE ID=" & y
						rs.Open strSQL, cn, 0, 3
						strRecipe = strRecipe & rs(0) & "<BR>" & VbCrLf
						rs.Close
					End If
				Next
			Next
		End If
		'------------------------------------------------------------------------
		Response.Write("<HR>" & strName & "(" & intID & ")<BR>")
		Response.Write(strRecipe & "<BR>")
		Response.Write(strDescription & "<BR>")
	Else
		Response.Write("<HR>DELETED (" & intID & ")<BR>")
	End If
End Sub

%>

