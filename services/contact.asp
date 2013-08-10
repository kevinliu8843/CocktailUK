<%
Option Explicit
strTitle="Contact Us"
Dim strMailTo
If Request("mail") <> "" Then
	strMailTo = Request("mail")
Else
	strMailTo = "theteam"
End If
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->

  <H2 align="left">Contact us</H2>
<FORM action="/mail/mail.asp" method="POST" name=form1 onSubmit="return checkFields()">
<div align="center">
<%If Request("fail") <> "" Then%>
	<p align="center"><b><font color="#FF0000"><%=Request("fail")%></font></b></p>
<%End If%>

<%If Request("message") <> "" Then%>
	<p align="center"><b><font color="#FF0000"><%=Request("message")%></font></b></p>
<%End If%>

  <center>
<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border: 1px solid #612B83" bordercolor="#111111" id="AutoNumber1">
  <TR>
    <TD width="100%"><IMG border="0" src="../images/main_menus/contactus.gif"></TD>
  </TR>
  <TR>
    <TD width="100%" background="../images/grad_write_purple_large.gif" style="padding: 8px;"> 
    <P align="left"> 
<B>Your name</B> 
<BR><INPUT TYPE=TEXT NAME=name SIZE=39 value="<%=Session("firstName") & " " & Session("lastName")%>"><BR> 
<B>E-mail address</B> 
<BR><INPUT TYPE=TEXT NAME=replyemail SIZE=39 value="<%=Session("email")%>"><br>
<B>Comments:</B> 
<BR><TEXTAREA WRAP=VIRTUAL NAME="comments" 
 COLS=31 ROWS=5></TEXTAREA>  

<P <%If Request("fail") <> "" Then%>style="color: red;"<%End If%> align="left"><b>Please enter the code below:</b><br> 
<!--#include virtual="/includes/CAPTCHA/CAPTCHA_form_inc.asp" --></p>

<P align="center"> 
<INPUT type="submit" value="Send comments &raquo;" name="B1" class="button" ><BR>
&nbsp;</TD>
  </TR>
</TABLE>
  </center>
</div>
<INPUT type="hidden" name="thankyou" value="/thankYou.asp"> 
<input type="hidden" name="mailto" value="<%=strMailTo%>">
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
	else if (document.form1.comments.value == "")
	{
		alert("Please enter some comments!")
		document.form1.comments.focus()
		return false
	}
	else
	{
		return true
	}
}
</SCRIPT>
<!--#include virtual="/includes/footer.asp" -->