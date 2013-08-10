<%Option Explicit%>
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/variables.asp" -->
<%
strTitle = "Cocktails By Ingredient"
Dim cn, i, strIngredientList, objDict, arrIngredients 
%>
<!--#include virtual="/includes/header.asp" -->
<H2>Find a cocktail by ingredient</H2>
<%
set cn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB
If Session("logged") Then
	strSQL = "SELECT ingredientID FROM UsrIng WHERE memID=" & Session("ID")
	rs.Open strSQL, cn, 0, 3
	Set objDict = CreateObject("Scripting.Dictionary")
	objDict.RemoveAll()
	While Not rs.EOF
		objDict.Add CStr(rs("ingredientID")),CStr(rs("ingredientID"))
		rs.movenext
	Wend
	rs.close
End If
%>
<P align="center">Please click the ingredient:<%If Session("logged") Then%> (your ingredients are in bold)<%End If%>
<P><Font color=red><I><%=Request.QueryString("error")%></I></Font>

<TABLE border="0" cellpadding="2" cellspacing="0" bordercolor="#000000" width="100%">
<% 
For i=0 To g_intNumIngredientTypes
	Response.Write("<tr><td colspan=""3"" align=center><B><BIG><FONT color=""#612b83"">" & Capitalise(g_aryIngredientType(i)) & "</FONT></BIG></B></td></tr>")
	Call DisplayIngredients(g_aryIngredientTypeID(i), objDict)
Next
%>

</TABLE>
<%
cn.Close
Set cn = Nothing
Set rs = Nothing

Sub DisplayIngredients(intType, objDict)
	Dim intPos, aryRows, intNumRows, i, j, strBGColor
	Dim strBG1, strBG2
	strBG1 = "#f0f0f0"
	strBG2 = "#ebe2f7"
	
	strSQL = "SELECT ID, name FROM ingredients WHERE Status=1 And Type=" & intType & " ORDER BY name"
	
	rs.Open strSQL, cn, 0, 3
	aryRows = rs.GetRows()
	rs.Close

	intNumRows = Int(UBound(aryRows, 2) / 3)
	For i=0 to intNumRows
		Response.Write("<tr valign=""top"">")
		For j=0 to 2
			intPos = i+(j*(intNumRows+1))
			
			strBGColor = strBG1
			If intPos<=UBound(aryRows, 2) Then
				Response.Write("<td bgcolor=""#F0F0F0"">")
				If Session("logged") then
					If objDict.Exists(CStr(aryRows(0, intPos))) Then
						Response.write("<B>")
					End If
				End If
				Response.Write("<A href=""/cocktails/containing.asp?ingredient=" & aryRows(0, intPos) & """>")
				Response.Write(capitalise( aryRows(1, intPos) ) & "</a></nobr>" & VbCrLf)
				If Session("logged") then
					If objDict.exists(aryRows(0, intPos)) Then
						Response.write("</B>")
					End If
				End If
				Response.Write("</td>")
			Else
				Response.Write("<td bgcolor=""" & strBGColor & """>&nbsp;</td>")
			End If
			
		Next
		Response.Write("</tr>")
	Next
End Sub
Call DrawSearchCocktailArea()%><!--#include virtual="/includes/footer.asp" -->