<%
Option Explicit

Dim cn, intID, strType, strDrinkName, blnNoDesc, intDrinkID, Mail, strMailTo, strBody, strTextArea
Dim strImgSrc, FSO, strMessage, strDescription, objFilter
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" -->
<!--#include virtual="/includes/forum/cforum.inc" -->
<!--#include virtual="/includes/CAPTCHA/CAPTCHA_process_form.asp" -->
<%
intID = Replace(Request("id"), ",", "")
If NOT IsNumeric(intID) OR intID = "" Then
	Response.redirect("/")
End If

strTextArea = "Possibly include what you think of it, where it can be purchased in the UK, even what you consider to be a good brand."

set cn = Server.CreateObject("ADODB.Connection")
set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDBMod

If Request.QueryString("delete") <> "" Then
	cn.Execute("DELETE from drink_desc WHERE id=" & Int(Request("delete")))
	Response.Write("Deleted!")
End If

'Update fields
If Request("update") = "true" AND Request("description") <> strTextArea Then
	strSQL = "SELECT name FROM ingredients WHERE id=" & intID
	rs.Open strSQL, cn, 0, 3
	strDrinkName = rs("name")
	rs.Close
	strDescription = strIntoDB(Replace(Request("description"), VbCrLf, "<BR>"))
	Set objfilter = New CForum
	strDescription = objFilter.FilterALLSwearWords(strDescription, False)
	Set objFilter = nothing

	If NOT IsSpam(strDescription) AND blnCAPTCHAcodeCorrect Then
		strSQL = "EXECUTE CUK_NEWINGDESCRIPTION @n='" & strIntoDB(Request("name")) & "', @d='" & strDescription & "', @id=" & intID
		cn.Execute(strSQL)
		strMessage="Thank you, your submission has been recieved, we just need to eyeball it before letting it on to the site!"
	Else
		strMessage="Please enter the correct code"
	End If
End If

'Read Fields
strSQL = "SELECT name FROM Ingredients WHERE id=" & intID
rs.Open strSQL, cn, 0, 3
If Not rs.EOF Then
	strDrinkName = rs("name")
	Set FSO = CreateObject("Scripting.FileSystemObject")
	If FSO.FileExists(Server.MapPath("/images/ingredients/" & intID & ".jpg" )) then
		strImgSrc="/images/ingredients/" & intID & ".jpg"
	Else
		strImgSrc="/images/ingredients/no_image.gif"
	End If
	Set FSO = nothing
End If
rs.Close

strTitle = "Ingredient Description - " & strDrinkName
%>
<!--#include virtual="/includes/header.asp" -->
<SCRIPT language="javascript">
function checkFields()
{
	if (document.drink.name.value == "")
	{
		alert("Please enter a name for yourself")
		document.drink.name.focus()
		return false
	}
	else if (document.drink.description.value == "")
	{
		alert("Please enter some comments!")
		document.drink.description.focus()
		return false
	}
	else
	{
		return true
	}
}

function clearArea()
{
	strString = document.drink.description.value
	if (strString.indexOf("<%=strTextArea%>") >= 0)
	{
		document.drink.description.value = ""
	}
}
function showSubmit(){
	document.getElementById('submitForm').style.display='block'
}
</script>
<h2><%=strTitle%></h2>
<table border="0" cellpadding="4" width="100%" style="border-collapse: collapse">
  <tr>
    <td valign="top" colspan="2">
    <b><a href="javascript:history.go(-1)">&lt; &lt; Back</a></b></td>
  </tr>
  <tr>
    <td valign="top">
          <table cellspacing="0" cellpadding="0" width="100%" border="0" id="table1">
            <tr>
              <td class="arrowblock" align="left" width="1%" nowrap>
              <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
              <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;IMAGE</b></td>
            </tr>
          </table>
    <P align="center">
    <IMG SRC="<%=strImgSrc%>" align="center" style="border: 1px solid #000000">
    </td>
    <td valign="top">

          <table cellspacing="0" cellpadding="0" width="100%" border="0" id="table2">
            <tr>
              <td class="arrowblock" align="left" width="1%" nowrap>
              <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
              <td class="baselightred" width="99%"><b class="contentHeader">
				&nbsp;DESCRIPTIONS</b></td>
            </tr>
          </table>

<%If strMessage <> "" then%>
	<P align="center"><FONT color="#FF0000"><%=strMessage%></FONT></P>
<%End If%>
<%
strSQL = "SELECT * from drink_desc WHERE status=1 AND drink_id=" & intID
rs.Open strSQL, cn, 0, 3
If NOT rs.EOF Then
	blnNoDesc = False
Else
	blnNoDesc = True
End If

If blnNoDesc Then
%>
<H5 align="center"><SPAN style="font-weight: 400">
<B>There are currently no user descriptions for <%=strDrinkName%></B></SPAN></H5>
<%
Else
Do While Not rs.EOF%> 
	<P><B><%=rs("name")%>:</b><%If bIsAdmin Then%>&nbsp;<A href="ingredient_Description.asp?delete=<%=rs("id")%>&type=<%=strType%>&id=<%=intID%>">delete</A><%End If%><BR><%=Trim(rs("description"))%></b>
	<%
	rs.MoveNext
Loop
rs.Close
cn.Close

Set cn = Nothing
Set rs = Nothing
%>
<%
End If
%>    </td>
  </tr>
  <tr>
    <td valign="top" colspan="2">
          <table cellspacing="0" cellpadding="0" width="100%" border="0" id="table3">
            <tr>
              <td class="arrowblock" align="left" width="1%" nowrap>
              <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
              <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;ADD 
				YOUR OWN DESCRIPTION</b></td>
            </tr>
          </table>

<FORM name="drink" action="ingredient_description.asp" onSubmit="return checkFields()">

<input type="hidden" name="type" value="<%=strType%>">
<input type="hidden" name="id" value="<%=intID%>">
<input type="hidden" name="update" value="true">
<div align="center" ID="submitForm">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" id="table4">
    <tr>
    <td>
    <blockquote>
      <b>Your Name:</b><BR>
      <input type="text" name="name" size="37" value="<%=Session("name")%>" style="width: 90%"></p>
    </blockquote>
    </td>
    </tr>
    <tr>
    <td>
    <blockquote>
      <b>Comments/description of <%=strDrinkName%>:</b><BR>
      <textarea rows="5" name="description" cols="28" maxlength="255" onFocus="clearArea()" style="width: 90%"><%=strTextArea%></textarea></blockquote>
    </td>
    </tr>
    <tr>
    <td>
    <blockquote>
      <P <%If Request("fail") <> "" Then%>style="color: red;"<%End If%> style="width:400"><b>Please enter the code below:</b><br> 
      <!--#include virtual="/includes/CAPTCHA/CAPTCHA_form_inc.asp" --></p>
    </blockquote>
    </td>
    </tr>
    <tr>
    <td valign="top">
      <p align="center">
    <INPUT type="submit" value="Add Description &gt; &gt;" name="B1" class="button" ></p>
      </td>
    </tr>
    <tr>
      <td></td>
    </tr>
  </table>
  </center>
</div>
</FORM>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->