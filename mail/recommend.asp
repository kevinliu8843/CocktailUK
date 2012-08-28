<%
Option Explicit
strTitle="Recommend a Friend"
%>

<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
	  <H2 align="left">Recommend Cocktail : UK to a friend</H2>
<form action="mail.asp?type=friend" method="POST" onSubmit="return checkFields()" name=form1> 

<div align="center">
  <center> 

<TABLE BORDER=0 CELLPADDING=0 cellspacing="0" style="border: 1px solid #612B83">  
	<TR>
	  <TD> 
      <img border="0" src="../images/main_menus/recommend.gif" width="283" height="42"></TD>
	</TR>
	<TR>
	  <TD background="../images/grad_write_purple_large.gif"> 
      <p align="center">
	    Thank you for recommending us.
	    <P align="center">&nbsp;<B>Your name</B><BR>  
        <INPUT NAME=name SIZE=30 value="<%=Session("firstName") & " " & Session("lastName")%>"><BR><B>Your e-mail address</B> 
	    <BR>&nbsp;<INPUT NAME=replyemail SIZE=30 value="<%=Session("email")%>"><B><br>
		Your friends' name</B><BR><INPUT NAME=friendname SIZE=30><BR><B>Your 
		friends' e-mail address</B> 
	    <BR><INPUT NAME=friendemail SIZE=30><B><br>
		Optional comments<BR>
        </B>
        <TEXTAREA rows="4" name="comments" cols="23"></TEXTAREA>
	    </P>
	    <P align="center">
        <INPUT type="checkbox" name="confirm" value="ON" checked id="fp2"><label for="fp2">&nbsp;Send me a confirmation e-mail</label><INPUT type="hidden" name="html" value="ON" checked id="fp1"><P align="center">
        <INPUT  type=submit name="submit" value="Submit &gt; &gt;" class="button"> 
        <br>
&nbsp;</P>
	  </TD>
	</TR>
	</TABLE>

  </center>
</div>

</FORM>
<SCRIPT language="Javascript">
function checkFields()
{
	if (document.form1.name.value == "")
	{
		alert("Please enter a name for yourself")
		document.form1.name.focus()
		return false
	}
	else if (document.form1.replyemail.value == "")
	{
		alert("Please enter an email for yourself")
		document.form1.replyemail.focus()
		return false
	}
	else if (document.form1.friendname.value == "")
	{
		alert("Please enter a name for your friend")
		document.form1.friendname.focus()
		return false
	}
	else if (document.form1.friendemail.value == "")
	{
		alert("Please enter an email for your friend")
		document.form1.friendemail.focus()
		return false
	}
	else
	{
		return true
	}
}
</SCRIPT>

<!--#include virtual="/includes/footer.asp" -->