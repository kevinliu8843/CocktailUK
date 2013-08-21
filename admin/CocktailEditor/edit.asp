<%
Option Explicit

Dim cn, intID, i, j, strAction
Dim strName, strIngredients, arySplit, intMeasure, intIngredient, strDescription, strUser, intType
Dim aryIngredients

CONST FONT_SMALL = "<font face=""Verdana, Arial"" size=""1"">"

%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" -->
<%

set cn			= Server.CreateObject("ADODB.Connection")
Set rs			= Server.CreateObject("ADODB.RecordSet")
cn.Open strDBMod


strAction		= Request("action")
intID			= Request("ID")
strName			= Replace(Request("Name"), VbCrLf, "")
strDescription	= Replace(Request("Description"), VbCrLf, "<br>")
strUser			= Replace(Request("User"), VbCrLf, "")
intType			= Int(Request("Type"))


Call GenerateJavascript()

If strAction = "update" Then 
	Call Update()
Else
	' i.e. don't load if we've just updated, since we know the field values
	If intID <> "" Then Call Load()
End If
Call ShowForm()

cn.Close
set cn			= Nothing
Set rs			= Nothing

Sub ShowForm()
	Response.Write("<table>")
	Response.Write("<form action=""edit.asp"" method=""post"">")
	Response.Write("<tr>")
	Response.Write("<td colspan=""2"" valign=""top"">")
	Response.Write("<input type=""hidden"" name=""ID"" value=""" & intID & """>")
	Response.Write("<input type=""submit"" name=""action"" value=""update""><BR>")
	Response.Write("</td>")
	Response.Write("</tr>")

	Response.Write("<tr>")
	Response.Write("<td valign=""top"">")

	Response.Write(FONT_SMALL & "Drink [" & intID & "]<br></font>")
	Response.Write("<input type=""text"" name=""name"" value=""" & Server.HTMLEncode(strName) & """><BR>")
	Response.Write(FONT_SMALL & "User<br></font>")
	Response.Write("<input type=""text"" name=""user"" value=""" & Server.HTMLEncode(strUser) & """><BR>")

	Response.Write(FONT_SMALL & "Type<br></font>")
	Response.Write("<SELECT name=""Type"">")
	Call AddSelectOption("1", "Cocktail", CStr(intType))
	Call AddSelectOption("2", "Shooter", CStr(intType))
	Call AddSelectOption("5", "Cocktail (non-alcoholic)", CStr(intType))
	Call AddSelectOption("6", "Shooter (non-alcoholic)", CStr(intType))
	Response.Write("</SELECT><BR>")
	
	Response.Write(FONT_SMALL & "Type<br></font>")

	Response.Write(FONT_SMALL & "Description<br></font>")
	Response.Write("<textarea cols=""30"" rows=""5"" name=""description"">" & Replace(strDescription, "<br>", VbCrLf) & "</textarea>")
	Response.Write("</td>")

	Response.Write("<td valign=""top"">")
	Response.Write(FONT_SMALL & "Ingredients (Check box to delete row)<br></font>")
	Call ShowIngredientSelector()
	Response.Write("</td>")

	Response.Write("</tr>")
	Response.Write("</form>")
	Response.Write("</table>")
End Sub

Sub ShowIngredientSelector()
	Response.Write("<table>")
	aryIngredients = GetRecipeArray(rs, cn, intID)
	If IsArray(aryIngredients) Then
		For i=0 to UBound(aryIngredients,2)+1
			If i<=UBound(aryIngredients,2) Then
				intMeasure		= aryIngredients(0,i)
				intIngredient	= aryIngredients(2,i) 
			Else
				intMeasure		= -1
				intIngredient	= -1
				Response.Write("<tr><td colspan=""3""><hr></td></tr>")
			End If
	
			Response.Write("<tr>")
			If i<=UBound(aryIngredients,2) Then
				Response.Write("<td>" & i+1 & "</td>")
			Else
				Response.Write("<td>New</td>")
			End If
	
			Response.Write("<td><script language=""javascript"">outputMeasuresList(""MeasureRow" & i & """, """ & intMeasure & """);</script></td>")
				Response.Write("<td><script language=""javascript"">outputIngredientsList(""IngredientRow" & i & """, """ & intIngredient & """);</script>")
	
			If i<=UBound(aryIngredients,2) Then
				Response.Write("<input type=""checkbox"" name=""DeleteRow" & i & """>")
			End If
	
			Response.Write("</td><tr>")
		Next
		Response.Write("</table>")
	Else
		intMeasure		= -1
		intIngredient	= -1
		Response.Write("<tr>")
		Response.Write("<td>New</td>")
		Response.Write("<td><script language=""javascript"">outputMeasuresList(""MeasureRow0"", """ & intMeasure & """);</script></td>")
		Response.Write("<td><script language=""javascript"">outputIngredientsList(""IngredientRow0"", """ & intIngredient & """);</script>")
		Response.Write("</td><tr>")
		
	End If
End Sub

Sub Update()
	Dim i

	strIngredients = ""		'Leave this as a global var.
	i=0
	cn.execute("DELETE FROM CocktailIng WHERE CocktailID="&intID)
	Do 
		If Request("MeasureRow" & i) <> "" AND Request("IngredientRow" & i) <> "" AND Request("DeleteRow" & i) <> "on" Then
			intMeasure	= Int(Request("MeasureRow" & i))
			intIngredient	= Int(Request("IngredientRow" & i))
			cn.execute("INSERT INTO CocktailIng (CocktailID, ingredientID, measureID) VALUES ("&intID&", "&intIngredient&", "&intMeasure&")")
		Else
			Exit Do
		End If
		i=i+1
	Loop
	
	strSQL = "UPDATE Cocktail SET "
	strSQL = strSQL & "ReIndex = 1, "
	strSQL = strSQL & "Type = " & intType & ", "
	strSQL = strSQL & "Name = '" & Replace(strName, "'", "''") & "', "
	strSQL = strSQL & "usr = '" & Replace(strUser, "'", "''") & "', "
	strSQL = strSQL & "Description = '" & Replace(strDescription, "'", "''") & "' "
	strSQL = strSQL & "WHERE ID=" & intID

	cn.Execute(strSQL)
	Response.Write("<FONT color=red>UPDATED</Font><BR>")
End Sub

Sub Load()
	strSQL = "SELECT * FROM Cocktail WHERE ID=" & intID
	rs.Open strSQL, cn, 0, 3
	If Not rs.EOF Then
		strName			= Trim(rs("Name"))
		strDescription	= Trim(rs("Description"))
		strUser			= Trim(rs("usr") & "")
		intType			= Int(rs("Type"))
	End If
	rs.Close
	aryIngredients = GetRecipeArray(rs, cn, intID)
End Sub

Public Sub AddSelectOption(strValue, strText, strSelectedValue)
	' Adds an entry into an HTML <SELECTion> box
	Response.Write("<option ")
	If Trim(CStr(strValue)) = Trim(CStr(strSelectedValue)) Then
		Response.Write("selected ")
	End If
	Response.Write("value=""" & strValue & """>" & strText)
End Sub

Sub GenerateJavascript()
	Dim aryMeasures, aryIngredients, i
	strSQL = "SELECT ID, Name FROM Measure ORDER BY Name"
	rs.Open strSQL, cn, 0, 3
	aryMeasures = rs.GetRows
	rs.Close
	strSQL = "SELECT ID, Name FROM Ingredients ORDER BY Name"
	rs.Open strSQL, cn, 0, 3
	aryIngredients = rs.GetRows
	rs.Close
%>
<script language="javascript">
function outputMeasuresList(strName, strSelected)
{
	var i
	var a = new Array(<%=UBound(aryMeasures, 2)+1%>)
<% 
	For i=0 To UBound(aryMeasures, 2) 
		Response.Write("a[" & i & "] = new Array(""" & aryMeasures(0, i) & """, """ & aryMeasures(1, i) & """)" & VbCrLf )
	Next
%>
	document.writeln("<select name=\"" + strName + "\">");

	document.writeln("<option selected value=\"-1\">[NOT SELECTED]" );
	for(i=0; i<<%=UBound(aryMeasures, 2)+1%>; i++)
	{
		if(strSelected == a[i][0])
			document.writeln("<option selected value=\"" + a[i][0] + "\">" + a[i][1] );
		else
			document.writeln("<option value=\"" + a[i][0] + "\">" + a[i][1] );
	}
	document.writeln("</select>");
}

function outputIngredientsList(strName, strSelected)
{
	var i
	var a = new Array(<%=UBound(aryIngredients, 2)+1%>)
<% 
	For i=0 To UBound(aryIngredients, 2) 
		Response.Write("a[" & i & "] = new Array(""" & aryIngredients(0, i) & """, """ & aryIngredients(1, i) & """)" & VbCrLf)
	Next
%>
	document.writeln("<select name=\"" + strName + "\">");

	document.writeln("<option selected value=\"-1\">[NOT SELECTED]" );
	for(i=0; i<<%=UBound(aryIngredients, 2)+1%>; i++)
	{
		if(strSelected == a[i][0])
			document.writeln("<option selected value=\"" + a[i][0] + "\">" + a[i][1] );
		else
			document.writeln("<option value=\"" + a[i][0] + "\">" + a[i][1] );
	}
	document.writeln("</select>");
}
</script>
<%
End Sub
%>