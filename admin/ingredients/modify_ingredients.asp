<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
Dim conn, i, j, intIndex, strSubmitButton, strTextBox, strAction
Dim aryRows, intNumRows, intPos, strName, strColor, bActive, intStatus, bAlcoholic, intAlcoholic
Dim intIngredientType, strDelete, strMove, intNumRecord, aryRows2, aryID, intNumIng

strTitle = "Modify Ingredients"
%>
<!--#include virtual="/includes/header.asp" -->

<h2>Modify Ingredients</h2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
<FORM name="ingredients" method="get" action="modify_ingredients.asp"> 
<SCRIPT Language="JavaScript">
function checkDel(){
	if (confirm("Are you sure you wish to delete these ingredients?")){
		Id = ""
		for (i=0; i<document.ingredients.todelete.length; i++){
			if (document.ingredients.todelete[i].checked){
				Id = Id + "|" + document.ingredients.todelete[i].value
			}
		}
		strHref="modify_ingredients.asp?ID="+Id+"&delete=true"
		if (confirm("Delete associated recipes too?"))
			strHref = strHref + "&recipes=true"
		location.href = strHref
	}
}
function checkMove(Id){
	if (confirm("Delete ingredient and change recipes to reflect change of ingredient?"))
	location.href="modify_ingredients.asp?ID="+Id+"&change=true&id2="+ingredients.change.options[ingredients.change.selectedIndex].value
}
</SCRIPT>
<%
strAction = Request("action")

Set conn	= Server.CreateObject("ADODB.Connection")
Set rs		= Server.CreateObject("ADODB.Recordset")
conn.Open strDBMod
intIngredientType = 0
For i = 0 To g_intNumIngredientTypes
	If Request("name" & g_aryIngredientTypeID(i)) <> "" then
		intIngredientType = Int(g_aryIngredientTypeID(i))
		Exit For
	End If
Next

strName = Replace(Request("Name" & intIngredientType ), "'", "''")
bActive = Request("active"&intIngredientType ) = "ON"
If bActive = True Then
	intStatus = 1
Else
	intStatus = 0
End If
bAlcoholic = Request("alcohol"&intIngredientType ) = "ON"
If bAlcoholic = True Then
	intAlcoholic = 1
Else
	intAlcoholic = 0
End If

If strAction = "Edit" AND intIngredientType > 0 AND request("delete") = "" Then
	'UPDATE Mode
	strSQL = "UPDATE ingredients Set name='" & strName & "', type=" & Request("type" & intIngredientType) & ", Status=" & intStatus & ", alcohol=" & intAlcoholic & " WHERE ID=" & Request("ID")
	conn.Execute(strSQL)
	response.write("<CENTER><FONT color=red>Ingredient updated</FONT></CENTER>")
elseif strAction = "New" AND intIngredientType > 0 Then
	'INSERT Mode
	strSQL = "INSERT into ingredients (name, type, Status, alcohol) VALUES('" & strName & "', " & intIngredientType & ", 1, " & intAlcoholic & ")"
	conn.Execute(strSQL)
	response.write("<CENTER><FONT color=red>Ingredient added</FONT></CENTER>")
elseif request("ID") <> "" AND request("delete") <> "" Then
	aryID = Split(request("ID"), "|")
	'DELETE Mode
	For i=1 To UBound (aryID)
		strSQL = "DELETE from ingredients WHERE ID=" & aryID(i)
		conn.Execute(strSQL)
		response.write("<CENTER><FONT color=red>Ingredient deleted</FONT></CENTER>")
		if (Request("recipes") = "true") then
			strSQL = "UPDATE cocktail set reindex=1, status=2 WHERE ID IN (SELECT DISTINCT cocktail.ID from cocktail,CocktailIng WHERE cocktail.ID = CocktailIng.cocktailID AND CocktailIng.IngredientID="&aryID(i)&")"
			conn.Execute(strSQL)
			response.write("<CENTER><FONT color=red>Associated recipes deleted</FONT></CENTER>")
		end if
	Next
elseif Request("ID") <> "" AND request("change") <> "" then
	strSQL = "DELETE from ingredients WHERE ID=" & Request("id")
	conn.Execute(strSQL)
	strSQL = "UPDATE CocktailIng Set ingredientID="&Request("id2")&" WHERE ingredientID="&Request("id")
	conn.execute(strSQL)
	response.write("<CENTER><FONT color=red>Ingredient deleted and associated recipes ("&intNumRecord&") changed</FONT></CENTER>")
End If

%>

<table width="100%" cellspacing="0" cellpadding="2">
<% For intIndex=0 To g_intNumIngredientTypes %>
	<tr>
<%
	If Request("edit") = g_aryIngredientTypeID(intIndex) & "" Then
    	strSubmitButton = "Edit"
		strSQL = "Select name, Status, alcohol from ingredients WHERE ID=" & Request("ID")
		rs.Open strSQL, conn, 0, 3
		strTextBox	= rs("name")
		bActive		= rs("Status") = "1"
		bAlcoholic  = rs("alcohol") = "1"

		Response.Write("<INPUT type=""hidden"" name=""ID"" value=""" & Request("ID") & """/>")
		strDelete = "&nbsp;<INPUT type=""button"" name=""delete"" value=""Del"" onClick=""checkDel(" & Request("ID") & ")"" style=""color: #FFFFFF; font-weight: bold; background-color: #612b83"">"
		strMove   = "<BR><INPUT type=""button"" name=""move"" value=""Change to:"" onClick=""checkMove(" & Request("ID") & ")"" style=""color: #FFFFFF; font-weight: bold; background-color: #612b83""><BR><SELECT id=""change"">"
		rs.close
		rs.Open "SELECT name, id from ingredients order by name", conn
		while not rs.EOF
			strMove = strMove & "<OPTION value="""&rs("id")&""" "
			If rs("id")&""=Request("ID") then
				strMove = strMove & "selected"
			End If
			strMove = strMove & ">"&rs("name")&"</OPTION>"
			rs.movenext
		wend
		strMove = strMove & "</SELECT>"
		rs.Close 
	Else
		strSubmitButton	= "New"
		strTextBox		= ""
		strDelete		= ""
		strMove 		= ""
   End If
	'bAlcoholic = Request("alcohol") = "ON"
	'bActive = Request("active") = "ON"   
%>
    <td colspan="1" width="20%" align="left"> <B><%=g_aryIngredientType(intIndex)%></B> </td>
	<td>
    <INPUT type=text name="name<%=g_aryIngredientTypeID(intIndex)%>" value="<%=strTextBox%>" size="20"><BR>
	<SELECT name="type<%=g_aryIngredientTypeID(intIndex)%>">
	<%
	For j = 0 To g_intNumIngredientTypes
	%>
		<OPTION value="<%=g_aryIngredientTypeID(j)%>" <%If j=intIndex then%>selected<%End If%>><%=g_aryIngredientType(j)%></OPTION>
	<%
	Next
	%>
	</SELECT><br>

	Active <input name="active<%=g_aryIngredientTypeID(intIndex)%>" type="checkbox" <%If bActive = True Then Response.Write("checked")%> value="ON"> Alcoholic <input name="alcohol<%=g_aryIngredientTypeID(intIndex)%>" type="checkbox" <%If bAlcoholic = True Then Response.Write("checked")%> value="ON"></td>
	<td>
	<INPUT type="submit" name="action" value="<%=strSubmitButton%>" class="button"><%=strDelete%><%=strMove%>
	</td>
	</tr>
<%
	' Output the ingredients for this type...
	strSQL = "Select Ingredients.ID, Ingredients.name, Ingredients.Status, Ingredients.alcohol FROM Ingredients WHERE Type=" & g_aryIngredientTypeID(intIndex) & " ORDER BY name"
	rs.Open strSQL, conn, 0, 3
	aryRows = rs.GetRows()
	rs.Close
	intNumRows = Int(UBound(aryRows, 2) / 3)
	For i=0 to intNumRows
		Response.Write("<tr>")
		For j=0 to 2
			intPos = i+(j*(intNumRows+1))
			
			If intPos<=UBound(aryRows, 2) Then

				Response.Write("<td bgcolor=""#F0F0F0"" valign=top width=""33%"">")
				strColor = "#000000"
				If Int(aryRows(2, intPos)) <> 1 Then 
					strColor = "#FF0000"
				ElseIf Int(aryRows(3, intPos)) <> 1 Then
					strColor = "#990077"
				End If

				'Get number of drinks this ingredient is in...
				'strSQL = "Select count(*) FROM cocktail WHERE ingredients LIKE'%-" & aryRows(0,intPos) & ",%'"
				'rs.Open strSQL, conn, 0, 3
				'intNumIng = rs(0)
				'rs.Close
				
				Response.Write("<TABLE cellpadding=0 cellspacing=0><TR><TD valign=top><INPUT type=""checkbox"" name=""todelete"" value="""&aryRows(0,intPos)&"""")
				If Request("ID")&"" = aryRows(0,intPos)&"" then
					response.write("checked")
				End If
				Response.write("></TD><TD><A style=""color: " & strColor & ";"" href=""?edit=" & g_aryIngredientTypeID(intIndex) & "&id=" & aryRows(0, intPos) & """>")
				Response.Write(capitalise( aryRows(1, intPos) ) & "</a></nobr>" & VbCrLf)

				Response.Write("</TD></TR></TABLE></td>")
			Else
				Response.Write("<td bgcolor=""#F0F0F0"">&nbsp;</td>")
			End If
			
		Next
		Response.Write("</tr>")
	Next
Next 
%>

</table>
</FORM>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->