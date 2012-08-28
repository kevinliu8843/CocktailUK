<%
Option Explicit
Server.ScriptTimeout = 10000

CONST NO_MATCH			= 0
CONST INGREDIENT_MATCH	= 1
CONST BOTH_MATCH		= 2

Dim cn, rsCheck, aryAllCocktails, i, j, bHighlight
%>
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/variables.asp" -->

<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
Set rsCheck = Server.CreateObject("ADODB.RecordSet")
rsCheck.Fields.Append "IngredientID",  3
rsCheck.Fields.Append "MeasureID",  3
rsCheck.Open

cn.Open strDBMod

strSQL = "SELECT ID, Ingredients, Name FROM Cocktail WHERE Status<>2 ORDER BY ID"
rs.Open strSQL, cn, 0, 3
aryAllCocktails = rs.GetRows
rs.Close


' Sort all ingredients --------------------------------------------
For i=0 To UBound(aryAllCocktails,2)
	aryAllCocktails(1, i) = SortList(aryAllCocktails(1, i))
Next


' Check all drinks --------------------------------------------------------
bHighlight = True
Response.Write("<table width=""100%"">")
Response.Write("<tr>")
Response.Write("<td>Drink</td>")
Response.Write("<td>Duplicates</td>")
Response.Write("<td>Action</td>")

Response.Write("</tr>")

For i=0 To UBound(aryAllCocktails,2)
	Call CheckSimilar(i)
Next

Response.Write("</table>")

cn.Close
Set cn = Nothing
Set rs = Nothing
Set rsCheck = Nothing


Sub CheckSimilar(intIndex)
	Dim intID, strIngredients, intResult, i, strSimilar, bFirst

	strSimilar = ""
	bFirst = True

	intID			= aryAllCocktails(0, intIndex)
	strIngredients	= aryAllCocktails(1, intIndex)

	For i=0 To UBound(aryAllCocktails,2)
		If Int(aryAllCocktails(0, i)) > Int(intID) Then

			intResult = CheckIngredients(strIngredients, aryAllCocktails(1, i))

			If intResult > NO_MATCH Then
				If strSimilar <> "" Then strSimilar = strSimilar & "|"
				strSimilar = strSimilar & aryAllCocktails(0, i) & "," & intResult

				If bFirst Then
					If bHighlight Then
						Response.Write("<tr bgcolor=""#F0F0F0"">")
					Else
						Response.Write("<tr>")
					End If
					bHighlight = Not bHighlight

					Response.Write("<td width=""1""><nobr>" & aryAllCocktails(2, intIndex) & " (" & aryAllCocktails(0, intIndex) & ") </td><td width=""1""><nobr>")
					bFirst = False
				End If

				If intResult = BOTH_MATCH Then 
					Response.Write("<font color=""red"">")
				Else
					Response.Write("<font color=""green"">")
				End If
				Response.Write(aryAllCocktails(2, i) & " (" & aryAllCocktails(0, i) & ")</font><br>")
			End If
		End IF
	Next

	If strSimilar <> "" Then
		Response.Write("</td><td valign=""top"">")
		Response.Write("<a target=""editduplicates"" href=""editduplicates.asp?id=" & intID & "&similar=" & Server.URLEncode(strSimilar) & """>Edit</a>")
		Response.Write("</td></tr>")
	End If
End Sub

Function CheckIngredients(strLeft, strRight)
	Dim arySplitLeft, arySplitRight, i, bFound
	Dim intTotalCount, intIngredientCount, intMeasureCount

	If strLeft = "" And strRight = "" Then
		CheckIngredients = BOTH_MATCH
		Exit Function
	ElseIf strLeft = "" Or strRight = "" Then
		CheckIngredients = NO_MATCH
		Exit Function
	End If
	
	arySplitLeft	= Split(strLeft, ",")
	arySplitRight	= Split(strRight, ",")

	If UBound(arySplitLeft) <> UBound(arySplitRight) Then
		CheckIngredients = NO_MATCH
		Exit Function
	Else
		CheckIngredients = BOTH_MATCH

		For i=0 To UBound(arySplitLeft)-1
			If Split(arySplitLeft(i), "-")(1) <> Split(arySplitRight(i), "-")(1) Then
				CheckIngredients = NO_MATCH
				Exit Function
			ElseIf Split(arySplitLeft(i), "-")(0) <> Split(arySplitRight(i), "-")(0) Then
				CheckIngredients = INGREDIENT_MATCH
			End If
		Next
	End If
End Function

Function SortList(strList)
	Dim aryList, i, aryFields
	aryFields = Array("IngredientID", "MeasureID")

	If Instr(strList, ",") Then
		aryList = Split(strList, ",")
	Else
		SortList = ""
		Exit Function
	End If
	
	For i=0 to UBound(aryList)-1
		rsCheck.AddNew aryFields, Array(Trim(Split(aryList(i), "-")(1)), Trim(Split(aryList(i), "-")(0)) )
	Next
	rsCheck.Sort = "IngredientID, MeasureID"
	
	Do While Not rsCheck.EOF
		SortList = SortList & Trim(rsCheck("MeasureID")) & "-" & Trim(rsCheck("IngredientID")) & ","

		rsCheck.Delete
		rsCheck.MoveNext
	Loop
End Function

%>

