<%
Option Explicit
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
strTitle="My Web Bar"
Dim cn, strIngredientList, i, strBG1, strBG2, aryIngs

If Not Session("logged") Then
	Response.Redirect("/db/member/loginout.asp")
End If

'Update bar here...
If Request("submit_form") = "true" Then
	set cn = Server.CreateObject("ADODB.Connection")
	cn.Open strDBMod
	cn.Execute("DELETE FROM usrIng WHERE memID="&Session("ID"))
	If Request("Ingredients") <> "" Then
		aryIngs = Split(Request("Ingredients"), ",")
		For i=0 TO UBound(aryIngs)
			If aryIngs(i) <> "" AND IsNumeric(aryIngs(i)) Then
				strSQL = "INSERT INTO UsrIng(memID, ingredientID) VALUES("&Session("ID")&", "&aryIngs(i)&")"
				cn.Execute(strSQL)	
			End If
		Next
		Session("recipes") = ""
	End If
	response.Redirect("/db/member/loginOut.asp?message=Your%20bar%20has%20been%20updated")
End If

strBG1 = "#f0f0f0"
strBG2 = "#e0e0e0"

set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDBMod
%>
<!--#include virtual="/includes/header.asp" --> 
<SCRIPT Language="javascript">
function changeBG(intTbl){
	var oTbl = document.getElementById("t"+intTbl)
	if (oTbl.style.backgroundColor=="<%=strBG1%>")
		oTbl.style.backgroundColor="<%=strBG2%>"
	else
		oTbl.style.backgroundColor="<%=strBG1%>" 
}
</SCRIPT>
<h2><%=Session("firstName")%>'s bar</FONT></h2>

<P align="center">Select the ingredients that you have and click save.</P>
<FORM method="POST" action="selectIngredients.asp">
<CENTER><INPUT TYPE="reset" VALUE="Undo Changes" class="button" > 
    <INPUT type="submit" value="Save list &gt; &gt;" name="B1" class="button" ><BR>&nbsp;</CENTER>
	<table cellpadding="0" cellspacing="0">
<%
	For i=0 To g_intNumIngredientTypes
		Response.Write("<tr><td colspan=""3"" align=center><B><BIG><FONT color=""#612b83"">" & Capitalise(g_aryIngredientType(i)) & "</FONT></BIG></B></td></tr>")
		
		Call DisplayIngredients(g_aryIngredientTypeID(i))
	Next
%>
	</table>
<BR>
<CENTER><INPUT TYPE="reset" VALUE="Undo Changes" class="button"612b83" > 
    <INPUT type="submit" value="Save list &gt; &gt;" name="B2" class="button" ><BR>&nbsp;</CENTER>
<INPUT type="hidden" name="submit_form" value="true">
</FORM>
<%

cn.Close
Set cn = Nothing
Set rs = Nothing

Sub DisplayIngredients(intType)
	Dim intPos, strList, aryRows, intNumRows, i, j, strBGColor
	
	strSQL = "SELECT 1 AS usr, ID, name FROM ingredients WHERE Status=1 And ID IN (SELECT IngredientID FROM UsrIng WHERE memID="&Session("ID")&") AND Type=" & intType
	strSQL = strSQL & " UNION "
	strSQL = strSQL & "SELECT 0 As usr, ID, name FROM ingredients WHERE Status=1 And ID NOT IN (SELECT IngredientID FROM UsrIng WHERE memID="&Session("ID")&") AND Type=" & intType & " ORDER BY usr  DESC, name"
	'Else
	'	strSQL = "SELECT 0 As usr, ID, name FROM ingredients WHERE Status=1 And Type=" & intType & " ORDER BY name"
	'End If
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
				If aryRows(0, intPos) = "1" Then strBGColor = strBG2

				Response.Write("<td ID=""t" & aryRows(1, intPos) & """ style=""background-color:" & strBGColor & ";"">")
				Response.Write("<TABLE cellspacing=0 cellpadding=0><TR><TD width=""5"" valign=top><INPUT type=""checkbox"" name=""ingredients"" onClick=""changeBG(" & aryRows(1, intPos) & ")"" value=""" & aryRows(1, intPos) & """ ID=""" & aryRows(1, intPos) & """ ")
				If aryRows(0, intPos) = "1" Then
					Response.Write("checked>")
				Else
					Response.Write(">")
				End If
				
				Response.Write("</TD><TD><LABEL for=""" & aryRows(1, intPos) & """>" & capitalise( aryRows(2, intPos) ) & "</LABEL></TD></TR></TABLE>"&VbCrLf)
				Response.Write("</td>")
			Else
				Response.Write("<td bgcolor=""" & strBGColor & """>&nbsp;</td>")
			End If
			
		Next
		Response.Write("</tr>")
	Next
End Sub
%><!--#include virtual="/includes/footer.asp" -->