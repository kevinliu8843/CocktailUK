<%
Option Explicit
strTitle = "Forgotten Details"
%>

<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->

<H2>Forgotten your login details?</H2>

<form action="mail.asp?type=details" method="POST" onSubmit="return checkFields()" name=form1> 
<TABLE BORDER=0 CELLPADDING=7>  
	<TR>
	  <TD align="center"> 
        <%If Request.QueryString("error") <> "" Then%>
        <P><I><FONT color="#FF0000">The email address that you supplied is not the email address you supplied when you registered</FONT></I>
        </P>
        <p align="left">
        <%End If%>
        Please enter your email address that you entered when you registered? If
        you have forgotten which email address you used you can <A href="/services/contact.asp">contact
        us</A> with your first and last name and I will send you your details.</p>
        <P><B>Enter your e-mail address?</B> 
	    <BR><INPUT NAME=email SIZE=40></P>
	  </TD>
	</TR>
	<TR>
	  <TD ALIGN=CENTER> 
		<P>
        <INPUT  type=submit name="submit" value="Get my login details &gt; &gt;" class="button" ></P>
      </TD> 
	</TR>
</TABLE>

</FORM>
<SCRIPT language="Javascript">
function checkFields()
{
	if (document.form1.email.value == "")
	{
		alert("Please enter an email address for yourself")
		document.form1.email.focus()
		return false
	}
	else
	{
		return true
	}
}
</SCRIPT>

<!--#include virtual="/includes/footer.asp" -->