<%
Option Explicit
strTitle="Email a Page"
Dim strReferer
strReferer = Request.ServerVariables("HTTP_REFERER")
If strReferer = "" Then
	Response.Redirect("/")
End If
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->

<FORM action="mail.asp?type=page" method="POST" name=form1 onSubmit="return checkFields()">
<INPUT TYPE="hidden" name="url" value="<%=strReferer%>">
<div align="center">
  <center>
<TABLE BORDER=0 CELLPADDING=7 style="border-collapse: collapse" bordercolor="#111111" cellspacing="0">  
<TR><TD> 
<H2>Email page to a friend</H2>
</TD></TR>  
<TR><TD> 
<P align="center"><B>Page:</B> <%=Request("pagetitle")%><BR>
(<A href="<%=strReferer%>"><%=strReferer%></A>)</TD></TR>  
<TR><TD align="center"> 
<i>This facility sends a short text email to a friend of yours with the location
of the page that you just came from. (With some optional extra comments from you
too...)</i>
</TD></TR><TR><TD align="center"> 
<P align="center"><B>What's your name?</B> 
	    <BR><INPUT NAME=name SIZE=40 value="<%=Session("firstName") & " " & Session("lastName")%>"> 
	    <BR> 
	    <BR><B>Your e-mail address?</B> 
	    <BR><INPUT NAME=replyemail SIZE=40 value="<%=Session("email")%>">&nbsp;</P>
	    <P align="center"><B>What's your friends' name?</B> 
	    <BR><INPUT NAME=friendname SIZE=40> 
	    <BR> 
	    <BR><B>Your friends' e-mail address?</B> 
	    <BR><INPUT NAME=friendemail SIZE=40>&nbsp;

<P><B>Your (optional) comments to them:</B> 
<BR><TEXTAREA WRAP=VIRTUAL NAME="comments" COLS=30 ROWS=5></TEXTAREA> 

</TD></TR><TR><TD ALIGN=CENTER> 
 <INPUT type=reset  name="reset" value="Clear form" class="button">
 <INPUT  type=submit name="submit" value="Submit" class="button">
  </TD> 
</TR></TABLE>
  </center>
</div>
<INPUT type="hidden" name="thankyou" value="/thankYou.asp"> 

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