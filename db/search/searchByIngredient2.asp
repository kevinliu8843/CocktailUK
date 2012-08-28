<%Option Explicit%>
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/variables.asp" -->
<%
strTitle = "Search By Ingredient"
Dim cn, i, strIngredientList, objDict, arrIngredients 
%>

<!--#include virtual="/includes/header.asp" -->

<script language="JavaScript">
var currentPanel;

function showPanel(panelNum) {
  //hide visible panel, show selected panel, 
  //set tab
  if (currentPanel != null) {
     hidePanel();
  }
  document.getElementById ('panel'+panelNum).style.display = 'block';
  currentPanel = panelNum;
  setState(panelNum);
}

function hidePanel() {
  //hide visible panel, unhilite tab
  document.getElementById('panel'+currentPanel).style.display = 'none';
  document.getElementById('tab'+currentPanel).style.backgroundColor = '#ffffff';
  document.getElementById('tab'+currentPanel).style.color = '#612b83';
}

function setState(tabNum) {
  if (tabNum==currentPanel) {
     document.getElementById('tab'+tabNum).style.backgroundColor = '#ddddff';
     document.getElementById('tab'+tabNum).style.color = 'red';
  }
  else {
     document.getElementById('tab'+tabNum).style.backgroundColor = '#ffffff';
     document.getElementById('tab'+tabNum).style.color = '#612b83';
  }
}
</script>

<H2>Find a drink by ingredient</H2>
<%
set cn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB
If Session("logged") Then
	strSQL = "SELECT ingredients FROM usr WHERE ingredients IS NOT NULL And ID=" & Session("ID")
	rs.Open strSQL, cn, 0, 3
	Set objDict = CreateObject("Scripting.Dictionary")
	objDict.RemoveAll()
	If Not rs.EOF Then
		strIngredientList = rs("ingredients")
		arrIngredients = Split(strIngredientList,",")
		For i=0 To UBound(arrIngredients)-1
			objDict.Add CStr(arrIngredients(i+1)),CStr(arrIngredients(i+1))
		Next
	End If
	rs.close
End If
%>
<P>Please click the ingredient:<%If Session("logged") Then%> (your ingredients are in bold)<%End If%>
<P><Font color=red><I><%=Request.QueryString("error")%></I></Font>
<table border="0" width="700" cellspacing="0" cellpadding="0" id="table1">
	<tr>
		<td nowrap>
			<ul class="tabnav">
			<%For i=0 To g_intNumIngredientTypes%>
				<li><a href="#" <%if i=0 then%>class="active" <%end if%> onClick="showPanel(<%=i+1%>);"><%=Capitalise(g_aryIngredientType(i))%></a></li>
			<%Next%>
			</ul>
		</td>
	</tr>
	<tr>
		<td>
<%
For i=0 To g_intNumIngredientTypes
	Response.write("<div id=""panel"&i+1&""" class=""panel""")
	If i<>0 Then
		Response.write(" style=""display: none;""")
	End If
	Response.write("><TABLE width=""100%"" cols=""3"">")
	Call DisplayIngredients(g_aryIngredientTypeID(i), objDict)
	Response.write("</TABLE></DIV>")
Next
%>
</td>
	</tr>
</table>
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
		Response.Write("<tr>")
		For j=0 to 2
			intPos = i+(j*(intNumRows+1))
			
			strBGColor = strBG1
			If intPos<=UBound(aryRows, 2) Then
				Response.Write("<td bgcolor=""#F0F0F0"" valign=top>")
				If Session("logged") then
					If objDict.Exists(CStr(aryRows(0, intPos))) Then
						Response.write("<B>")
					End If
				End If
				Response.Write("<A href=""/db/findCocktailContIng.asp?ingredient=" & aryRows(0, intPos) & """>")
				Response.Write(capitalise( aryRows(1, intPos) ) & "</a>" & VbCrLf)
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
%><!--#include virtual="/includes/footer.asp" -->