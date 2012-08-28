<%
'Option Explicit

Dim cn, intBased, strBased, strIngredients
Dim aryRecipe, strRecipe, strName, strDescription, strType, strUser, intType, intServes, i, strXXX
Dim aryDrink

strTitle = "Submit Drinks"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/admin_functions.asp" -->
<H2>Cocktails awaiting verification</h2>
<FORM action="default.asp?submit=true" method=post>
<%
If Request.QueryString ("duplicated") <> "" Then
	Response.Write "<P><FONT color=red>Cocktail already exists</FONT></P>"
End If

set cn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.Recordset")
cn.Open strDBMod

If Request.QueryString("delete") <> "" Then
	'**** **** Status 2 = Deleted - Check database design for details.... **** **** **** **** **** **** 
	'**** Don't need to reindex since this drink won't have been live already **** ****
	Set rs = cn.Execute("UPDATE Cocktail Set Status=2, ReIndex=0 WHERE ID=" & Int(Request.QueryString("delete")))
End If

If Request("submit") = "true" Then
	Set rs2 = cn.Execute ("SELECT ID,usr,name from Cocktail WHERE Status=0")
	While NOT rs2.EOF
		user = CStr(rs2("usr"))
		name = Capitalise(rs2("name"))
		strSQL = "UPDATE Cocktail Set Status=1, ReIndex=1, usr='" & replaceStuff(Left(user, InStr(1, user, ";")-1)) & "' WHERE ID=" & rs2("ID")
		cn.Execute(strSQL)
		call sendCocktailsubmitEmail(replaceStuff( CStr( name ) ), Right(user, Len(user)-InStr(1, user, ";")))
		rs2.movenext
	Wend
End If

Set rs2 = cn.Execute ("SELECT name, ID, dateadded from Cocktail WHERE Status=0 ORDER by name")

Do While NOT rs2.EOF
	'Response.Write "<A href=""viewCocktail.asp?ID="& rs("ID") &""">" & replaceStuffBack(rs("name")) & "</a> - (<A href=default.asp?delete="& rs("ID") &">Delete Recipe</a>), added: "&rs("dateadded")&"<BR>" & VbCrLf
'---------------------------------------------------------------------------------------------------------------
strName = GetAwaitingDrink(rs, cn, rs2("ID"), aryDrink)
strtitle = aryDrink(0)
i = i + 1
%>
  <TABLE border="0" cellpadding="0" cellspacing="10" width="100%">
    <TR>
      <TD valign="top" colspan="2"><h3><%=i%>) <%=aryDrink(0)%></h3></TD>
    </TR>
    <TR>
      <TD valign="top" colspan="2">Category: <%=aryDrink(7)%>, <%=aryDrink(8)%>, <%=aryDrink(9)%></TD>
    </TR>
    <TR>
      <TD valign="top">
          <table cellspacing="0" cellpadding="0" width="100%" border="0" id="table1">
            <tr>
              <td class="arrowblock" align="left" width="1%" nowrap>
              <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
              <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;INSTRUCTIONS</b></td>
            </tr>
          </table>
        <%=aryDrink(1)%><P>
        <b>Serves <%=aryDrink(3)%></TD>
      <TD valign="top">
          <table cellspacing="0" cellpadding="0" width="100%" border="0" id="table2">
            <tr>
              <td class="arrowblock" align="left" width="1%" nowrap>
              <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
              <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;INGREDIENTS</b></td>
            </tr>
          </table>
        <nobr><%=aryDrink(2)%></nobr>
      </TD>
    </TR>
    <TR>
      <TD valign="top" colspan="2">User details : <%=aryDrink(10)%></TD>
    </TR>
  </TABLE>
  <TABLE border="0" cellpadding="0" cellspacing="0" width="100%">
    <TR>
      <TD width="50%">
        <P align="center"><INPUT class="button" type="button" value="Delete" name="B3" onClick="location.href='default.asp?delete=<%=rs2("ID")%>'"></P>
      </TD>
      <TD width="50%">
        <P align="center"><INPUT class="button" type="button" value="Edit" name="B4" onClick="top.location.href='/admin/default.asp?goto=cocktaileditor/default.asp?ID=<%=rs2("ID")%>'"></P>
      </TD>
    </TR>
  </TABLE>
        <HR size="1" color="#000000">
<%
	rs2.MoveNext
Loop
rs2.close
cn.close
Set cn = nothing
Set rs = nothing
%>
<CENTER><INPUT class="button" type="submit" value="Submit these recipes &gt; &gt;" name="B4"></CENTER>
</FORM>
<!--#include virtual="/includes/footer.asp" -->

