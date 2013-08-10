<%
Option Explicit
%>
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" -->
<!--#include virtual="/includes/variables.asp" --><%
Dim cn, strName, strType, strDirections, intServes, intBased, strBased, strError, intStage
Dim aryIngredients, aryMeasures, i, bFound, strXXX, blnDuplicated, strID

ReDim aryIngredients(g_intNumIngredientTypes)
ReDim aryMeasures(g_intNumIngredientTypes)

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDBMod


' Determine Stage ----------------------------------------------------
intStage=0
For i=0 to 6
	If Request("S" & i) <> "" Then
		intStage = i
		Exit For
	End If
Next

' Pull in form data ----------------------------------------------------------
strName			= Capitalise(Request("name"))
strType			= Request("type")
strXXX			= Request("xxx")
strDirections	= Capitalise(Request("directions"))
intServes		= Request("serves")

For i=0 to g_intNumIngredientTypes
	aryIngredients(i)	= Request("IngredientType" & g_aryIngredientTypeID(i))
	aryMeasures(i)		= Request("IngredientMeasure" & g_aryIngredientTypeID(i))
Next
intBased = 0
If Request("based") <> "" Then 
	If IsNumeric(Request("based")) Then intBased = Int(Request("based"))
End If

If intBased <> 0 Then
	strSQL = "SELECT name FROM Ingredients WHERE ID=" & intBased
	rs.Open strSQL, cn, 0, 3
	strBased = Capitalise(rs("name"))
	rs.Close
End If


' Error handling ------------------------------------------------------------
If intStage = 2 Then
	'Check duplicate cocktail names...
	rs.open "SELECT count(*), ID from cocktail where name='"&strIntoDB(strName)&"' AND status=1 GROUP BY ID", cn, 0, 3
	blnDuplicated = False
	If NOT rs.EOF Then
		If rs(0) > 0 Then
			blnDuplicated = True
			strID = rs(1)
		End If
	End If
	rs.close
End If

If intStage > 1 And (strName = "" Or strType = "" Or blnduplicated) Then
	intStage	= 1
	If strName = ""	Then 
		strError = "Please enter a name for the recipe"
	ElseIf strType = "" Then
		strError = "Please choose a drink type"
	ElseIf blnDuplicated Then
		strError = "<BR>We already have a recipe with this name. Please check if it is the same recipe by viewing it <A HREF=""/db/viewCocktail.asp?ID="&strID&""">here</A>"
	End If
End If

If strError = "" And intStage > 2 Then
	bFound = False
	For i= 0 To g_intNumIngredientTypes
		If aryIngredients(i) <> "" Then
			bFound = True
			Exit For
		End If
	Next

	If bFound = False Then
		strError = "Please choose some ingredients"
		intStage=2
	End If
End If

If strError = "" And intStage > 4 And strDirections = "" Then
	intStage = 4
	strError = "Please enter directions for this " & strType
ElseIf strError = "" And intStage > 4 And intBased = 0 And aryIngredients(SPIRIT) <> "" Then
	intStage = 4
	strError = "Please select a base for this " & strType
End If
'-----------------------------------------------------------------------------
Select Case intStage
	Case 0 strTitle = "Submit A Recipe"
	Case 1 strTitle = "Stage 1/5"
	Case 2 strTitle = "Stage 2/5"
	Case 3 strTitle = "Stage 3/5"
	Case 4 strTitle = "Stage 4/5"
	Case 5 strTitle = "Stage 5/5"
	Case 6 strTitle = "Complete"
End Select
%>
<!--#include virtual="/includes/header.asp" -->
<h2>Submit a recipe</h2>
<TABLE cellpadding="5" border="0" width="100%"><TR><TD>
<%
Select Case intStage
	Case 0 Call DisplayStage0()
	Case 1 Call DisplayStage1()
	Case 2 Call DisplayStage2()
	Case 3 Call DisplayStage3()
	Case 4 Call DisplayStage4()
	Case 5 Call DisplayStage5()
	Case 6 Call DisplayStage6()
End Select

cn.Close
Set rs = Nothing
Set cn = Nothing
%> <%
Sub DisplayStage0()
	If Session("logged") Then
%>
<p><b>Thank you for contributing to Cocktail : UK.</b> </p>
<p>Five screens will follow, asking for various details about the cocktail. At any 
stage you can go back to the previous one if you make a mistake.</p>
<blockquote>
	<dl>
		<dt>Your name : </dt>
		<dd><%=Session("firstname") & " " & Session("lastname")%></dd>
		<dt>Your email address (not ever displayed) : </dt>
		<dd><%=Session("email")%></dd>
	</dl>
	<p><a href="editProfile.asp">Edit these details now...</a></p>
</blockquote>
<p align="right">Click next to continue...</p>
<p align="center"><font color="#FF0000">Please do not submit recipes that contain 
anything that is libellous, defamatory, obscene, abusive. Such submissions will 
be removed</font></p>
<form method="POST" action="submitCocktail.asp" name="form1">
	<p align="right"><% Call DisplayButton("Next &gt; &gt;", 1) %> </p>
</form>
<script language="javascript">
document.form1.S1.focus();
</script>
<%
	Else
%>
<p>Thank you for wanting to contribute to Cocktail : UK</p>
<p>In order for us to continue with the on-line submission of recipes, please either:
</p>
<ol>
	<li><a href="/account/login.asp">Login</a>, so we can identify and recognise 
	you as the contributor of the recipe</li>
	<li><a href="/account/register.asp">Create a new account</a></li>
</ol>
<p><%
	End If
End Sub

Sub DisplayStage1()
%> </p>
<h3>Stage 1/5</h3>
<p>Please enter the recipe name and drink type... </p>
<form method="POST" action="submitCocktail.asp" name="form1">
	<%
	For i=0 to g_intNumIngredientTypes
		Response.Write("<INPUT type=""hidden"" name=""" & g_aryIngredientType(i) & """ value=""" & aryIngredients(i) & """>")
	Next
%>
	<table border="0" cellpadding="0" cellspacing="5" width="100%">
		<tr>
			<td valign="top" bgcolor="#F0F0F0">&nbsp;Name:
			<input type="text" name="name" size="20" value="<%=Server.HTMLEncode(strName)%>">
			<%If strError <> "" Then%><font color="red"><i><%=strError%></i></font><%End If%></td>
		</tr>
		<tr>
			<td valign="top" bgcolor="#F0F0F0">&nbsp;Type:
			<input type="radio" value="cocktail" <%if strtype = "cocktail" then%>checked<%end if%> name="type" id="fp1"><label for="fp1">Cocktail</label> 
			OR
			<input type="radio" name="type" <%if strtype = "shooter" then%>checked<%end if%> value="shooter" id="fp2"><label for="fp2">Shooter</label>
			</td>
		</tr>
		<tr>
			<td valign="top" bgcolor="#F0F0F0">&nbsp;Is the drink XXX rated?
			<input type="radio" name="xxx" value="XXX Rated" id="fp4"><label for="fp4">Yes</label>
			<input type="radio" name="xxx" value="Not XXX Rated" checked id="fp3"><label for="fp3">No</label> <br>
&nbsp;i.e. because of a &quot;rude&quot; title etc...</td>
		</tr>
	</table>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td>
			<p align="left"><% Call DisplayButton("&lt; &lt; Back", 0) %></p>
			</td>
			<td>
			<p align="center"><% Call DisplayButton(" Start over ", 0) %></p>
			</td>
			<td>
			<p align="right"><% Call DisplayButton("Next &gt; &gt;", 2) %></p>
			</td>
		</tr>
	</table>
</form>
<script language="javascript">
	document.form1.name.focus()
</script>
<%
End Sub

Sub DisplayStage2()
	Dim sql, strSelected, intNum, i, j, aryChosenIngredients
	Dim intPos, strList, aryRows, intNumRows, intCount, intIndex, strBGColor
	Dim bFound
	Dim strBG1, strBG2

	strBG1 = "#f0f0f0"
	strBG2 = "#ebe2f7"
%>
<script language="javascript">
function changeBG(intTbl){
	var oTbl = document.getElementById("t"+intTbl)
	if (oTbl.style.backgroundColor=="<%=strBG1%>")
		oTbl.style.backgroundColor="<%=strBG2%>"
	else
		oTbl.style.backgroundColor="<%=strBG1%>"
}
</script>
<h3>Stage 2/5</h3>
<p>Please select the ingredients that are in the <%=strType%>.</p>
<form method="POST" action="submitCocktail.asp" name="form1">
	<input type="hidden" name="name" value="<%=Server.HTMLEncode(strName)%>">
	<input type="hidden" name="type" value="<%=Server.HTMLEncode(strType)%>">
	<input type="hidden" name="xxx" value="<%=Server.HTMLEncode(strXXX)%>">
	<table border="0" cellpadding="0" cellspacing="5" width="100%">
		<tr>
			<td valign="top">
			<h2><%=strName%> </h2>
			</td>
		</tr>
		<tr>
			<td valign="top">&nbsp;Type: <%=strType%> (<%=strXXX%>) <% If strError <> "" Then %><br>
			<font color="red"><i><%=strError%></i><%End If %> </font></td>
		</tr>
		<tr>
			<td valign="top"><% Call ShowSubTitle("INGREDIENTS") %> </td>
		</tr>
		<tr>
			<td valign="top" bgcolor="#F0F0F0"><nobr>
			<table cellpadding="0" cellspacing="0">
				<%
	For intIndex=0 To g_intNumIngredientTypes
		Response.Write("<tr><td bgcolor=""#FFFFFF"" colspan=""3"" align=center><B><BIG><FONT color=""#612b83"">" & Capitalise(g_aryIngredientType(intIndex)) & "</FONT></BIG></B></td></tr>")

		If aryIngredients(intIndex) <> "" Then
			If Instr(aryIngredients(intIndex), ",") Then
				aryChosenIngredients = Split(aryIngredients(intIndex), ",")
			Else
				Redim aryChosenIngredients(0)
				aryChosenIngredients(0) = aryIngredients(intIndex)
			End If
		Else
			Redim aryChosenIngredients(0)
		End If

		strSQL = "SELECT ID, name FROM ingredients WHERE Status=1 And Type=" & g_aryIngredientTypeID(intIndex) & " ORDER BY name"
		
		rs.Open strSQL, cn, 0, 3
		aryRows = rs.GetRows()
		rs.Close

		intNumRows = Int(UBound(aryRows, 2) / 3)
		For i=0 to intNumRows
			Response.Write("<tr>")
			For j=0 to 2
				intPos = i+(j*(intNumRows+1))
				
				strBGColor = strBG1
				If intPos<=UBound(aryRows, 2) Then
					bFound = False
					For intCount=0 To UBound(aryChosenIngredients)
						If CStr(aryRows(0, intPos)) = Trim(aryChosenIngredients(intCount)) Then
							bFound = True
							Exit For
						End If
					Next

					If bFound Then strBGColor = strBG2

					Response.Write("<td ID=""t" & aryRows(0, intPos) & """ style=""background-color:" & strBGColor & ";"" valign=""top"">")
					Response.Write("<TABLE cellspacing=0 cellpadding=0><TR><TD valign=""top""><INPUT type=""checkbox"" name=""IngredientType" & g_aryIngredientTypeID(intIndex) & """ onClick=""changeBG(" & aryRows(0, intPos) & ")"" value=""" & aryRows(0, intPos) & """ ID=""" & aryRows(0, intPos) & """ ")
					If bFound Then
						Response.Write("checked>")
					Else
						Response.Write(">")
					End If

					Response.Write("</TD><TD><LABEL for=""" & aryRows(0, intPos) & """>" & capitalise( aryRows(1, intPos) ) & "</LABEL>"&VbCrLf)
					Response.Write("</TD></TABLE></td>")
				Else
					Response.Write("<td bgcolor=""" & strBGColor & """>&nbsp;</td>")
				End If
				
			Next
			Response.Write("</tr>")
		Next

	Next
%>
			</table>
			</nobr></td>
		</tr>
	</table>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td>
			<p align="left"><% Call DisplayButton("&lt; &lt; Back", 1) %></p>
			</td>
			<td>
			<p align="center"><% Call DisplayButton(" Start over ", 0) %></p>
			</td>
			<td>
			<p align="right"><% Call DisplayButton("Next &gt; &gt;", 3) %></p>
			</td>
		</tr>
	</table>
</form>
<%
End Sub

Sub DisplayStage3()
	Dim aryAllMeasures, aryChosenMeasures, i, j, arySplit, intIngredientType, intChosenMeasure
%>
<h3>Stage 3/5</h3>
<p>Please select the amount of each ingredient in the <%=strType%>.</p>
<form method="POST" action="submitCocktail.asp" name="form1">
	<input type="hidden" name="name" value="<%=Server.HTMLEncode(strName)%>">
	<input type="hidden" name="type" value="<%=Server.HTMLEncode(strType)%>">
	<input type="hidden" name="xxx" value="<%=Server.HTMLEncode(strXXX)%>">
	<input type="hidden" name="based" value="<%=intBased%>">
	<input type="hidden" name="serves" value="<%=intServes%>">
	<input type="hidden" name="directions" value="<%=Replace( Server.HTMLEncode(strDirections), VbCrLf, "<BR>")%>">
	<%
	For i=0 to g_intNumIngredientTypes
		Response.Write("<INPUT type=""hidden"" name=""IngredientType" & g_aryIngredientTypeID(i) & """ value=""" & aryIngredients(i) & """>")
	Next
%>
	<table border="0" cellpadding="0" cellspacing="5" width="100%">
		<tr>
			<td valign="top">
			<h2><%=strName%> </h2>
			</td>
		</tr>
		<tr>
			<td valign="top">&nbsp;Type: <%=strType%> (<%=strXXX%>)</td>
		</tr>
		<tr>
			<td valign="top"><% Call ShowSubTitle("INGREDIENTS") %> </td>
		</tr>
		<tr>
			<td valign="top" bgcolor="#F0F0F0"><nobr>
			<table>
				<%
		' Get list of all measures 
		strSQL = "SELECT ID, Name FROM measure WHERE Status=1"
		rs.Open strSQL, cn, 0, 3
		aryAllMeasures = rs.GetRows()
		rs.Close

		'Get the names of the ingredients (and keep them in the same order as we selected them )
		For intIngredientType=0 to g_intNumIngredientTypes
			If aryIngredients(intIngredientType) <> "" Then
				If aryMeasures(intIngredientType) <> "" Then
					aryChosenMeasures = Split(aryMeasures(intIngredientType), ",")
				Else
					Redim aryChosenMeasures(0)
				End If

				strSQL = ""
				arySplit = Split(aryIngredients(intIngredientType), "," )
				For j=0 To UBound(arySplit)
					If strSQL <> "" Then strSQL = strSQL & " UNION "
					strSQL = strSQL & "SELECT ID, Name FROM Ingredients WHERE ID=" & arySplit(j)
				Next
				strSQL = strSQL & " ORDER BY Name"
				rs.Open strSQL, cn, 0, 3
				i=0
				Do While NOT rs.EOF
					Response.Write("<TR><TD><SELECT name=""IngredientMeasure" & g_aryIngredientTypeID(intIngredientType) & """>" )

					If i <= UBound(aryChosenMeasures) Then
						intChosenMeasure = aryChosenMeasures(i)
					Else
						intChosenMeasure = 0
					End If

					For j=0 To UBound(aryAllMeasures, 2)
						Call AddSelectOption(CStr(aryAllMeasures(0, j)), aryAllMeasures(1, j), CStr(intChosenMeasure))
					Next
					
					Response.Write("</SELECT></TD><TD>" & Capitalise(rs("Name")) & "</TD></TR>")
					rs.MoveNext
					i=i+1
				Loop
				rs.Close
			End If
		Next
%>
			</table>
			</nobr></td>
		</tr>
	</table>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td>
			<p align="left"><% Call DisplayButton("&lt; &lt; Back", 2) %></p>
			</td>
			<td>
			<p align="center"><% Call DisplayButton(" Start over ", 0) %></p>
			</td>
			<td>
			<p align="right"><% Call DisplayButton("Next &gt; &gt;", 4) %></p>
			</td>
		</tr>
	</table>
	<%
End Sub

Sub DisplayStage4()
	Dim strIngredient
%>
	<h3>Stage 4/5</h3>
	<p>Please input the directions on how to make the <%=strType%> and how many 
	it serves. Also specify what the <%=strType%> is based on.</p>
</form>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form5_Validator(theForm)
{

  if (theForm.directions.value.length > 500)
  {
    alert("Please enter at most 500 characters in the \"directions\" field.");
    theForm.directions.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="submitCocktail.asp" name="FrontPage_Form5" onsubmit="return FrontPage_Form5_Validator(this)" language="JavaScript">
	<input type="hidden" name="name" value="<%=Server.HTMLEncode(strName)%>">
	<input type="hidden" name="type" value="<%=Server.HTMLEncode(strType)%>">
	<input type="hidden" name="xxx" value="<%=Server.HTMLEncode(strXXX)%>"><%
	For i=0 to g_intNumIngredientTypes
		Response.Write("<INPUT type=""hidden"" name=""IngredientType" & g_aryIngredientTypeID(i) & """ value=""" & aryIngredients(i) & """>")
		Response.Write("<INPUT type=""hidden"" name=""IngredientMeasure" & g_aryIngredientTypeID(i) & """ value=""" & aryMeasures(i) & """>")
	Next
%>
	<table border="0" cellpadding="0" cellspacing="5" width="100%">
		<tr>
			<td valign="top" colspan="2">
			<h2><%=strName%> </h2>
			</td>
		</tr>
		<tr>
			<td valign="top" colspan="2">&nbsp;Type: <%=strType%> (<%=strXXX%>)
			<% If strError <> "" Then %><br>
			<font color="red"><i><%=strError%></i><%End If %> </font></td>
		</tr>
		<tr>
			<td valign="top"><% Call ShowSubTitle("DIRECTIONS") %> </td>
			<td valign="top"><% Call ShowSubTitle("INGREDIENTS") %> </td>
		</tr>
		<tr>
			<td valign="top" bgcolor="#F0F0F0" width="1%">
			<p align="center">
			<!--webbot bot="Validation" i-maximum-length="500" --><textarea rows="8" name="directions" cols="31"><%=Server.HTMLEncode(Replace(strDirections, "<BR>", VbCrLf))%></textarea></p>
			<p align="left"><b><font face="Arial">Serves&nbsp;&nbsp;&nbsp;&nbsp;
			<select size="1" name="serves"><% 
		For i=1 To 10
			Call AddSelectOption(CStr(i), CStr(i), CStr(intServes))
        Next
%></select>&nbsp;&nbsp;&nbsp; <br>
			<% 
If aryIngredients(SPIRIT) <> "" Then 
%> Based on <select size="1" name="based">
			<option value="0" selected>Please select a base...</option>
			<%
		strSQL = "SELECT ID, Name FROM Ingredients WHERE ID IN(" & aryIngredients(SPIRIT) & ") ORDER BY name"
		rs.Open strSQL, cn, 0, 3
		Do While NOT rs.EOF
			Call AddSelectOption(CStr(rs("ID")), Capitalise(rs("name")), CStr(intBased))
			rs.MoveNext
		Loop
		rs.Close
%></select></font></b> <%
End If
%> </p>
			</td>
			<td valign="top"><nobr><% Call ShowChosenMeasuresAndIngredients() %>
			</nobr></td>
		</tr>
	</table>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td>
			<p align="left"><% Call DisplayButton("&lt; &lt; Back", 3) %></p>
			</td>
			<td>
			<p align="center"><% Call DisplayButton(" Start over ", 0) %></p>
			</td>
			<td>
			<p align="right"><% Call DisplayButton("Next &gt; &gt;", 5) %></p>
			</td>
		</tr>
	</table>
	<%
End Sub

Sub DisplayStage5()
	Dim strIngredient
%>
	<h3>Stage 5/5</h3>
	<p>Please confirm the <%=strType%> details.</p>
</form>
<form method="POST" action="submitCocktail.asp" name="form1">
	<input type="hidden" name="name" value="<%=Server.HTMLEncode(strName)%>">
	<input type="hidden" name="type" value="<%=Server.HTMLEncode(strType)%>">
	<input type="hidden" name="xxx" value="<%=Server.HTMLEncode(strXXX)%>">
	<input type="hidden" name="directions" value="<%=Replace( Server.HTMLEncode(strDirections), VbCrLf, "<BR>")%>">
	<input type="hidden" name="serves" value="<%=intServes%>">
	<input type="hidden" name="based" value="<%=intBased%>"><%
	For i=0 to g_intNumIngredientTypes
		Response.Write("<INPUT type=""hidden"" name=""IngredientType" & g_aryIngredientTypeID(i) & """ value=""" & aryIngredients(i) & """>")
		Response.Write("<INPUT type=""hidden"" name=""IngredientMeasure" & g_aryIngredientTypeID(i) & """ value=""" & aryMeasures(i) & """>")
	Next
%>
	<table border="0" cellpadding="0" cellspacing="5" width="100%">
		<tr>
			<td valign="top" colspan="3">
			<h2><%=strName%> </h2>
			</td>
		</tr>
		<tr>
			<td valign="top" colspan="3">&nbsp;Type: <%=strType%> (<%=strXXX%>)</td>
		</tr>
		<tr>
			<td valign="top"><% 
	Call ShowSubTitle("DIRECTIONS") 
	Response.Write(Replace( strDirections, VbCrLf, "<BR>") )
%>
			<p><b><font face="Arial">Serves <%=intServes%></font> <% If intBased > 0 Then %><br>
			<font face="Arial">Based on <%=strBased%></font> </b><%End If%> </p>
			</td>
			<td valign="top">&nbsp;</td>
			<td valign="top"><%
	Call ShowSubTitle("INGREDIENTS")
	Call ShowChosenMeasuresAndIngredients()
%> </td>
		</tr>
	</table>
	<table border="0" cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td>
			<p align="left"><% Call DisplayButton("&lt; &lt; Back", 4) %></p>
			</td>
			<td>
			<p align="center"><% Call DisplayButton(" Start over ", 0) %></p>
			</td>
			<td>
			<p align="right"><% Call DisplayButton("Submit cocktail recipe", 6) %></p>
			</td>
		</tr>
	</table>
	<%
End Sub

Sub DisplayStage6()
	Dim intType, bHasErrors, i, strAllMeasures, intID
	Dim strSpirit, strLiquor, strMixer, strOther, strIngredients

	If strType = "cocktail" Then
		intType = 1
	Else
		intType = 2
	End If
	
	If strXXX = "XXX Rated" then
		intType = intType + 8
	End If
	
	'Add cocktail to db table ----------------------------------------------------
	strDirections	= replaceStuff(strDirections)

	strIngredients = SortIngredientIDs()
	For i=0 to g_intNumIngredientTypes
		If aryMeasures(i) <> "" Then 
			strAllMeasures = strAllMeasures & aryMeasures(i) & ","
		End If
	Next

	' See if we have a spare "slot" to use (instead of creating a new entry) ---------------
	strSQL = "SELECT TOP 1 ID FROM Cocktail WHERE Status=2 And ReIndex=0 ORDER BY ID"
	rs.Open strSQL, cn, 0, 3
	If Not rs.EOF Then
		intID = Int(rs("ID"))
	Else
		intID = -1
	End If
	rs.Close
	'---------------------------------------------------------------------------------------

	If intID = -1 Then
		' Create NEW
		strSQL = "INSERT INTO cocktail (name, description, type, Status, ReIndex, serves, based, dateadded, usr ) VALUES ("
		strSQL = strSQL & "'" & replaceStuff( CStr( strName ) ) & "', "
		strSQL = strSQL & "'" & CStr( strDirections ) & "', "
		strSQL = strSQL & intType & ", "
		strSQL = strSQL & "0, "
		strSQL = strSQL & "0, "
		strSQL = strSQL & CLng(intServes) & ", "
		strSQL = strSQL & Int(intBased) & ", "
		strSQL = strSQL & "GETDATE(), "
		strSQL = strSQL & "'" & CStr( Session("firstname") & " " & Session("lastname") & ";" & Session("email") ) & "'"
		
		strSQL = strSQL & ")"
	Else
		' Use an old deleted entry...
		cn.execute("DELETE FROM CocktailIng WHERE cocktailID="&intID)
		strSQL = "UPDATE ""cocktail"" SET "
		strSQL = strSQL & "Name = '" & replaceStuff( CStr( strName ) ) & "', "
		strSQL = strSQL & "Description = '" & CStr( strDirections ) & "', "
		strSQL = strSQL & "Type = " & intType & ", "
		strSQL = strSQL & "Rate = 0, "
		strSQL = strSQL & "Users = 0, "
		strSQL = strSQL & "Status = 0, "
		strSQL = strSQL & "Accessed = 0, "
		strSQL = strSQL & "ReIndex= 0, "
		strSQL = strSQL & "Serves = " & CLng(intServes) & ", "
		strSQL = strSQL & "Based = " & Int(intBased) & ", "
		strSQL = strSQL & "dateadded=GETDATE(), "
		strSQL = strSQL & "usr = '" & CStr( Session("firstname") & " " & Session("lastname") & ";" & Session("email") ) & "' "
		
		strSQL = strSQL & "WHERE ID = " & intID
	End If

	cn.Execute(strSQL)
	
	if intID = -1 Then
		strSQL = "SELECT TOP 1 ID from cocktail where name='"&replaceStuff( CStr( strName ) )&"' ORDER BY ID DESC"
		rs.open strSQL, cn, 0, 3
		If NOT rs.EOF Then
			intID = rs("ID")
		End If
	End If

	If (intID <> -1) Then
		aryIngredients= Split(strIngredients, ",")
		aryMeasures   = Split(strAllMeasures, ",")
		For i=0 To UBound(aryIngredients)
			If aryIngredients(i) <> "" AND aryMeasures(i) <> "" AND IsNumeric(aryIngredients(i)) AND IsNumeric(aryMeasures(i)) Then
				cn.execute("INSERT INTO CocktailIng (cocktailID, ingredientID, measureID) VALUES("&intID&", "&aryIngredients(i)&", "&aryMeasures(i)&")")
			End If
		Next
	End If

	bHasErrors = (err.number>0) OR (intID=-1)
	
	IF bHasErrors Then
		Response.Write "Database Errors Occured" & "<P>"
		for counter= 0 to conn.errors.count
			Response.Write "Error #" & conn.errors(counter).number & "<P>"
			Response.Write "Error desc. -> " & conn.errors(counter).description & "<P>"
		next
	End If

	'Finished adding cocktail to table --------------------------------------------------------

	If bHasErrors Then
%>
	<p><font color="red"><i>There was an error processing this page. Please
	<a href="/services/contact.asp">contact the webmaster</a></i></font></p>
	<%
	Else
%>
	<h2>Submission complete</h2>
	<p>The <%=strType%> has been submitted to the webmaster for review and action 
	will be taken shortly.</p>
	<p>Thank you for your time. Enjoy the <a href="/">rest of the site!</a></p>
	<table border="0" cellpadding="0" cellspacing="5" width="100%">
		<tr>
			<td valign="top" colspan="3">
			<h2><%=strName%> </h2>
			</td>
		</tr>
		<tr>
			<td valign="top" colspan="3">&nbsp;Type: <%=strType%> (<%=strXXX%>)</td>
		</tr>
		<tr>
			<td valign="top"><% 
		Call ShowSubTitle("DIRECTIONS") 
		Response.Write(Replace( strDirections, VbCrLf, "<BR>"))
%>
			<p><b><font face="Arial">Serves <%=intServes%></font> <%If intBased>0 Then %><br>
			<font face="Arial">Based on <%=strBased%></font> </b><%End If%> </p>
			</td>
			<td valign="top"></td>
			<td valign="top"><% Call ShowSubTitle("INGREDIENTS") %> <nobr><% Call ShowMeasuresAndIngredients(rs, cn, intID) %></nobr>
			</td>
		</tr>
	</table>
	<%
	End If
End Sub

Sub DisplayButton(strText, intStage)
	If intStage <> 6 Then
		Response.Write("<INPUT type=""submit"" class=""button"" value=""" & strText & """ name=""S" & intStage & """>")
	Else
		Response.Write("<INPUT type=""submit"" style=""color: #FFFFFF; font-weight: bold; background-color: red"" value=""" & strText & """ name=""S" & intStage & """>")
	End If
End Sub

Sub ShowMeasuresAndIngredients(rs, cn, intID)
	response.write GetRecipe(rs, cn, intID, true)
End Sub

Sub ShowChosenMeasuresAndIngredients()
	Dim strIngredients
	Dim i, j, arySplitIngredients, strAllMeasures, arySplitMeasure

	strIngredients = SortIngredientIDs()
	arySplitIngredients = Split(strIngredients, ",")
	For i=0 to g_intNumIngredientTypes
		If aryMeasures(i) <> "" Then 
			strAllMeasures = strAllMeasures & aryMeasures(i) & ","
		End If
	Next
	strAllMeasures = Left(strAllMeasures, Len(strAllMeasures)-1)

	arySplitMeasure		= Split(strAllMeasures, "," )

	strSQL = ""
	For i=0 To UBound(arySplitMeasure)
		If strSQL <> "" Then strSQL = strSQL & " UNION "
		strSQL = strSQL & "SELECT measure.name, Ingredients.Type, Ingredients.name FROM Measure, Ingredients WHERE Ingredients.ID=" & arySplitIngredients(i) & " AND measure.ID=" & arySplitMeasure(i)
	Next
	strSQL = strSQL & " ORDER BY Ingredients.Type, Ingredients.Name"
	rs.Open strSQL, cn, 0, 3
	Do While NOT rs.EOF
		Response.Write(rs(0) & " " & Capitalise(rs(2)) & "<BR>")
		rs.MoveNext
	Loop
	rs.Close
End Sub

Sub ShowSubTitle(strSubTitle)
%>
	<table cellspacing="0" cellpadding="0" width="100%" border="0">
		<tr>
			<td class="arrowblock" align="left" width="1%" nowrap><img height="16" src="/images/pixel.gif" width="16" border="0"></td>
			<td class="baselightred" width="93%"><b class="contentHeader">&nbsp;<%=strSubTitle%></b></td>
		</tr>
	</table>
<%
End Sub

Function SortIngredientIDs()
	' Ensure that the Ingredient IDs are in the same order as we entered
	' ( So they match with their respective Measure IDs)
	Dim i, j, arySplitIngredients, strIDs
	strIDs=""
	For i=0 to g_intNumIngredientTypes
		If aryIngredients(i) <> "" Then
			arySplitIngredients	= Split(aryIngredients(i), "," )
			j=0
			Do While j<=UBound(arySplitIngredients)
				If strIDs <> "" Then strIDs = strIDs & ","
				strIDs = strIDs & arySplitIngredients(j)
				j=j+1
			Loop
		End If
	Next

	strSQL = "SELECT ID FROM Ingredients WHERE ID IN (" & strIDs & ") Order BY Type, Name "
	
	rs.Open strSQL, cn, 0, 3
	Do While Not rs.EOF
		If SortIngredientIDs <> "" Then SortIngredientIDs = SortIngredientIDs & ","
		SortIngredientIDs = SortIngredientIDs & rs("ID")

		rs.MoveNext
	Loop
	rs.Close
End Function

Sub AddSelectOption(strValue, strText, strSelectedValue)
	' Adds an entry into an HTML <SELECTion> box
	Response.Write("<option ")
	If Trim(CStr(strValue)) = Trim(CStr(strSelectedValue)) Then
		Response.Write("selected ")
	End If
	Response.Write("value=""" & strValue & """>" & strText)
End Sub
%>
</TD></TR></TABLE><!--#include virtual="/includes/footer.asp" -->