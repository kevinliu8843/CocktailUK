<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/forum/cforum.inc" -->
<%
Dim intProd, intID, strName, cn, strReview, objFilter
intID = Request("ID")
intProd = Request("prodID")
If NOT IsNumeric(intID) OR NOT IsNumeric(intProd) Then
	Response.end
End If

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.open strDB
rs.open "SELECT name from cocktail WHERE ID=" & intID, cn, 0, 3
If NOT rs.EOF Then
	strName = replaceStuffBack(rs(0))
Else
	strName = ""
End If
rs.close
cn.close
set rs = nothing
set cn = nothing
%>
<html>

<head>
<title>Review a drink...</title>
<link rel="stylesheet" type="text/css" href="../style/style.css">
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="3" style="background-color: #FFFFFF">
<form method="GET" action="review.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
<table border="0" cellpadding="0" style="border-collapse: collapse" width="100%" id="table1">
	<tr>
		<td bgcolor="#FFFFFF">
		<a href="/"><img border="0" src="../images/cuk_03.jpg" width="85" height="85" align="middle"><img border="0" src="../images/cuk_07.gif" width="210" height="32" align="middle"></a></td>
	</tr>
	<tr>
		<td>
		<h3 align="center">Comment on the &quot;<%=strName%>&quot; drink </h3>
		</td>
	</tr>
</table>
<%
If intProd <> "" AND IsNumeric(intProd) Then
	'Submit form to db.
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.open strDBMod
	strReview = strIntoDB(Left(Request("review"),500))
	Set objfilter = New CForum
	strReview = objFilter.FilterALLSwearWords(strReview, False)
	Set objFilter = nothing
	If NOT IsSpam(strReview) Then
		strSQL = "EXECUTE CUK_NEWDRINKDESCRIPTION @n='" & strIntoDB(Request("name")) & "', @d='" & strReview & "', @id=" & intProd
		cn.execute(strSQL)
	End If
	cn.close
	Set cn = nothing
%>
<p align="center">Thank you for adding your comments.<br>
We just need to eyeball it before we display it on the site. <br>
<font color="#FFFFFF"><b><a href="#" onclick="self.close()">Close this window</a></b></font>
<%
	
ElseIf intID <> "" AND IsNumeric(intID) Then
	'Display form to accept review
	if strName <> "" Then
%>
	<div align="center">
		<center>
		<table border="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" width="375">
			<tr>
				<td align="right" valign="top" nowrap>
				<p align="left"><b>Your name:</b></p>
				</td>
			</tr>
			<tr>
				<td align="right" valign="top" nowrap>
				<input type="text" name="name" size="57" maxlength="50"></td>
			</tr>
			<tr>
				<td align="right" valign="top" nowrap>
				<p align="left"><b>Your comments of the drink: <font size="1">(less 
				than 500 characters please)</font>:</b></p>
				</td>
			</tr>
			<tr>
				<td align="right" valign="top" nowrap>
				<p align="center"><textarea rows="7" name="review" cols="43"></textarea></p>
				</td>
			</tr>
			<tr>
				<td align="right" valign="top">
				<p align="center">
    <INPUT type="submit" value="Add Comments &gt; &gt;" name="B1" class="button" >
				<br>
				<small>Please do not submit anything that is libellous, defamatory, 
				obscene, pornographic or abusive. Such posts will be removed.</small></p>
				</td>
			</tr>
		</table>
		</center>
	</div>
	<input type="hidden" name="prodID" value="<%=intID%>">
	<input type="hidden" name="ID" value="<%=intID%>">
</form>
<script language="JavaScript" type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.name.value == "")
  {
    alert("Please enter a value for the \"Name\" field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.name.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Name\" field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.name.value.length > 50)
  {
    alert("Please enter at most 50 characters in the \"Name\" field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.review.value == "")
  {
    alert("Please enter a value for the \"Review\" field.");
    theForm.review.focus();
    return (false);
  }

  if (theForm.review.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Review\" field.");
    theForm.review.focus();
    return (false);
  }

  if (theForm.review.value.length > 500)
  {
    alert("Please enter at most 500 characters in the \"Review\" field.");
    theForm.review.focus();
    return (false);
  }
  return (true);
}
//-->
</script>
<%Else%>
<script>
		self.close();
		</script>
<%End If%> <%Else%>
<script>
	self.close();
	</script>
<%End If%>

</body>

</html>