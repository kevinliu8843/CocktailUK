<% 
Option Explicit
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
Dim conn, strTableType, strDelete, strColor, i, bActive, intStatus, strTextBox, strSubmitButton
Dim strMove, strIng, intNumRecord, j, aryRows2, aryID, intNum
strTitle = "Modify Measures"
%>
<!--#include virtual="/includes/header.asp" -->

<h2>Modify Measures</h2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <tr>
    <td width="100%">
<FORM name="ingredients" method="get" action="modify_measures.asp"> 
<SCRIPT Language="JavaScript">
function checkDel(){
	if (confirm("Are you sure you wish to delete these measures?")){
		Id = ""
		for (i=0; i<document.ingredients.todelete.length; i++){
			if (document.ingredients.todelete[i].checked){
				Id = Id + "|" + document.ingredients.todelete[i].value
			}
		}
		strHref="modify_measures.asp?ID="+Id+"&delete=true"
		if (confirm("Delete associated recipes too?"))
			strHref = strHref + "&recipes=true"
		location.href = strHref
	}
}
function checkMove(Id){
	if (confirm("Delete ingredient and change recipes to reflect change of ingredient?"))
	location.href="modify_measures.asp?ID="+Id+"&change=true&id2="+ingredients.change.options[ingredients.change.selectedIndex].value
}
</SCRIPT>
<%
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open strDBMod
Set rs = Server.CreateObject("ADODB.Recordset")

if Request("measure") <> "" then
	strTableType = "measure"
end if

bActive = Request("active") = "on"
If bActive = True Then
	intStatus = 1
Else
	intStatus = 0
End If

If Request("ID") <> "" AND strTableType <> "" AND Request("delete") = "" Then
	'UPDATE Mode
	strSQL = "UPDATE measure Set name='" & Request(strTableType) & "', Status=" & intStatus & " WHERE ID=" & Request("ID")
	conn.Execute(strSQL)
	response.write("<CENTER><FONT color=red>Measure updated</FONT></CENTER>")
elseif Request("ID") = "" AND strTableType <> "" Then
	'INSERT Mode
	strSQL = "INSERT into measure (name, Status) VALUES('" & Replace(Request(strTableType), "'", "''") & "', 1)"
	conn.Execute(strSQL)
	response.write("<CENTER><FONT color=red>Measure added</FONT></CENTER>")
elseif Request("ID") <> "" AND Request("delete") <> "" Then
	aryID = Split(request("ID"), "|")
	'DELETE Mode
	For i=1 To UBound (aryID)
		strSQL = "DELETE from measure WHERE ID=" & aryID(i)
		conn.Execute(strSQL)
		response.write("<CENTER><FONT color=red>Measure deleted</FONT></CENTER>")
		if (Request("recipes") = "true") then
			strSQL = "UPDATE cocktail set reindex=1, status=2 WHERE ID IN (SELECT DISTINCT cocktail.ID from cocktail,CocktailIng WHERE cocktail.ID = CocktailIng.cocktailID AND CocktailIng.MeasureID="&aryID(i)&")"
			conn.Execute(strSQL)
			response.write("<CENTER><FONT color=red>Associated recipes deleted</FONT></CENTER>")
		end if
	Next
elseif Request("ID") <> "" AND request("change") <> "" then
	strSQL = "DELETE from measure WHERE ID=" & Request("id")
	conn.Execute(strSQL)
	strSQL = "UPDATE CocktailIng Set measureID="&Request("id2")&" WHERE measureID="&Request("id")
	conn.execute(strSQL)
	response.write("<CENTER><FONT color=red>Measure deleted and associated recipes changed</FONT></CENTER>")
End If
i=0
%>
<table border="1" width="100%" bordercolor="#000080" cellspacing="0" cellpadding="1" style="border-collapse: collapse;">
  <tr>
    <td width="20%" align="center"><B>Measures</B>&nbsp;</td>
  </tr>
  <tr>
<%
	If Request("edit") <> "measure" Then
		strSubmitButton = "New"
		strTextBox		= ""
		strDelete		= ""
		strMove			= ""
	 Else
		strSubmitButton = "Edit"
		strSQL = "SELECT name, Status FROM measure WHERE ID=" & Request("ID")
		rs.Open strSQL, conn, 0, 3
		strTextBox		= rs("name")
		bActive			= rs("Status") = "1"

		Response.Write("<INPUT type=""hidden"" name=""ID"" value=""" & Request("ID") & """>")
		strDelete = "&nbsp;<INPUT type=""button"" name=""delete"" value=""Del"" onClick=""checkDel()"" style=""color: #FFFFFF; font-weight: bold; background-color: #612b83"">"
		strMove   = "<BR><INPUT type=""button"" name=""move"" value=""Change to:"" onClick=""checkMove(" & Request("ID") & ")"" style=""color: #FFFFFF; font-weight: bold; background-color: #612b83""><BR><SELECT id=""change"">"
		rs.close
		rs.Open "SELECT name, id from measure WHERE status=1 order by name", conn
		while not rs.EOF
			strMove = strMove & "<OPTION value="""&rs("id")&""">"&rs("name")&"</OPTION>"
			rs.movenext
		wend
		strMove = strMove & "</SELECT>"
		rs.Close 
	End If
%>
    <td align="center">
    <INPUT type=text name="measure" value="<%=strTextBox%>" size="20"><BR>
<% 
	If Request("edit") = "measure" Then 
%>
	Active <input name="active" type="checkbox" <%If bActive = True Then Response.Write("checked")%> value="ON">
<%
	End If
%>
    <INPUT type="submit" value="<%=strSubmitButton%>" class="button"><%=strDelete%><%=strMove%>
	</td>
  </tr>
  <tr>
<%	
	If Request("edit") = "" Then
		strSQL = "SELECT COUNT(measure.name) AS intNum, measure.name, measure.ID, measure.status FROM CocktailIng INNER JOIN measure ON CocktailIng.measureID = measure.ID GROUP BY measure.name, measure.ID, measure.Status ORDER BY status ASC, name"
		rs.Open strSQL, conn, 0, 3
%>
    <td width="20%" valign="top">
<%
	 Do While Not rs.EOF
		If Int(rs("Status")) = 1 Then 
			strColor = "#000000"
		Else
			strColor = "#FF0000"
		End If
%>
    <INPUT type="checkbox" name="todelete" value="<%=rs("ID")%>" <%If Request("ID")&""=rs("ID")&"" then%>checked<%End If%>><A style="color:<%=strColor%>;" HREF="?edit=measure&id=<%=rs("ID")%>"><%=rs("name")%> (<%=rs("intNum")%>)</A><BR>
<%
			rs.movenext
	  Loop
	  rs.Close
  End If
%>
  </td>
  </tr>
</table>
</FORM>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->