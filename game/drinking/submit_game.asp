<%Option Explicit%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
Dim i, strName, strDirections, intType, cn, objFilter, blnDup, strRW

If Request("submit_form") = "true" Then
	'Check for dupliactes first
	blnDup = False
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open strDB
	strSQL = "SELECT count(*) FROM drinkinggame WHERE title LIKE '"&strTitle&"'"
	Set rs = cn.execute(strSQL)
	If rs(0) > 0 Then
		blnDup = True
	End If
	Set rs = nothing
	cn.close
	Set cn = Nothing
	
	If NOT blnDup Then
		If Request("html") <> "ON" Then
			strRW = "<BR>"
		End If
		Set objFilter = New CForum
		strName 		= strIntoDB(Request("name"))
		strName 		= objFilter.FilterALLSwearWords(strName , False)
		strTitle		= strIntoDB(Request("title"))
		strTitle		= objFilter.FilterALLSwearWords(strTitle, False)
		strDirections	= Replace(strIntoDB(Request("directions")), VbCrLf, strRW)
		strDirections 	= objFilter.FilterALLSwearWords(strDirections, False)
		Set objFilter = Nothing
		intType 		= Request("type")
		If NOT IsSpam(strDirections) Then
			Set cn = Server.CreateObject("ADODB.Connection")
			cn.Open strDB
			strSQL = "EXECUTE CUK_NEWDRINKINGGAME @n='"&strName&"', @title='"&strTitle&"', @t="&intType&", @d='"&strDirections&"'"
			cn.execute(strSQL)
			cn.close
			Set cn = Nothing
		End If
		Response.redirect("/game/drinking/default.asp?gameadded=true")	
	End if
End If

strTitle="Add game"
%>
<!--#include virtual="/includes/header.asp" -->
<H2>Add A Drinking Game</H2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber14">
  <tr>
    <td width="100%">
<P align="center">Here you can add a game you know and get credit for it too.</P>
<%If blnDup Then%>
<P align="center"><FONT color="#FF0000"><I>Sorry, that game already exists. Game not submitted...</I></FONT>
<%End If%> </P>
    <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.title.value == "")
  {
    alert("Please enter a value for the \"Game title\" field.");
    theForm.title.focus();
    return (false);
  }

  if (theForm.title.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Game title\" field.");
    theForm.title.focus();
    return (false);
  }

  if (theForm.title.value.length > 50)
  {
    alert("Please enter at most 50 characters in the \"Game title\" field.");
    theForm.title.focus();
    return (false);
  }

  if (theForm.type.selectedIndex == 0)
  {
    alert("The first \"Game type\" option is not a valid selection.  Please choose one of the other options.");
    theForm.type.focus();
    return (false);
  }

  if (theForm.directions.value == "")
  {
    alert("Please enter a value for the \"Game directions\" field.");
    theForm.directions.focus();
    return (false);
  }

  if (theForm.directions.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Game directions\" field.");
    theForm.directions.focus();
    return (false);
  }

  if (theForm.name.value == "")
  {
    alert("Please enter a value for the \"Your name\" field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.name.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Your name\" field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.name.value.length > 50)
  {
    alert("Please enter at most 50 characters in the \"Your name\" field.");
    theForm.name.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><FORM method="POST" action="submit_game.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
  <DIV align="center">
    <CENTER>
  <TABLE border="0" cellpadding="0" cellspacing="0" style="border-style:solid; border-width:1; border-collapse: collapse" bordercolor="#612B83" id="AutoNumber12" height="150">
    <TR>
      <TD height="10"><IMG border="0" src="../../images/main_menus/addagame.gif"></TD>
    </TR>
    <TR>
      <TD background="../../images/grad_write_purple_large.gif">
      <DIV align="center">
        <CENTER>
        <TABLE border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber13" cellpadding="0">
          <TR>
            <TD valign="top"><B>Game title</B></TD>
            <TD>
            <!--webbot bot="Validation" s-display-name="Game title" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="50" --><INPUT type="text" name="title" size="52" maxlength="50" value="<%=Request("title")%>"></TD>
          </TR>
          <TR>
            <TD valign="top"><B>Game type</B></TD>
            <TD>
            <!--webbot bot="Validation" s-display-name="Game type" b-disallow-first-item="TRUE" --><SELECT size="1" name="type">
            <OPTION value="-1">Please select...</OPTION>
            <%For i=0 to UBound(aryGames)%>
            <OPTION value="<%=i%>"><%=aryGames(i)%></OPTION>
            <%Next%>
            </SELECT></TD>
          </TR>
          <TR>
            <TD valign="top"><B>Game directions</B></TD>
            <TD>
            <!--webbot bot="Validation" s-display-name="Game directions" b-value-required="TRUE" i-minimum-length="1" --><TEXTAREA rows="12" name="directions" cols="39"><%=Request("directions")%></TEXTAREA></TD>
          </TR>
          <TR>
            <TD valign="top"><B>Your name</B></TD>
            <TD>
            <!--webbot bot="Validation" s-display-name="Your name" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="50" --><INPUT type="text" name="name" size="52" maxlength="50" value="<%=Request("name")%>"></TD>
          </TR>
          <%If bIsAdmin Then%>
          <TR>
            <TD valign="top"><B>HTML?</B></TD>
            <TD><INPUT type="checkbox" name="html" value="ON"></TD>
          </TR>
          <%End if%>
          <TR>
            <TD valign="top" colspan="2">
            <P align="center"><SMALL>Please do not submit anything that is libellous, defamatory, <BR>
            obscene, pornographic or abusive. Such games will be removed.<BR>
            <INPUT type="submit" value="Submit game" name="B1" style="color: #FFFFFF; font-weight: bold; background-color: #612b83; font-size:8pt; "></SMALL></TD>
          </TR>
        </TABLE>
        </CENTER>
      </DIV>
      </TD>
    </TR>
  </TABLE>
    <P align="center">Please note: we review submissions prior to acceptance. This is just so we can make sure all is ok with the submission...</P>
    </CENTER>
  </DIV>
  <INPUT type="hidden" name="submit_form" value="true">
</FORM>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->