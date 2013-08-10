<%
Option Explicit
strTitle="Edit Profile"
Dim cn, strPassword, blnNews, blnForum, strError
Dim strFirstName, strLastName, strEmail, strUname, strPass, strPass2, strExtra, blnOkay, intID
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<%
If NOT Session("logged") = True Then
	Response.Redirect("/default.asp")
End If

Set cn = Server.CreateObject("ADODB.Connection")
cn.Open strDBMod
blnOkay = False

If Request("submit_form") = "true" Then
	strFirstName 	= strIntoDB(Request("name"))
	strLastName 	= strIntoDB(Request("lname"))
	strEmail 		= strIntoDB(Request("email"))
	strUname 		= Replace(strIntoDB(Request("uname")), "%", "")
	strPass  		= Replace(strIntoDB(Request("pass")), "%", "")
	strPass2 		= Replace(strIntoDB(Request("pass2")), "%", "")
	blnNews  		= (Request("news")  = "ON")
	blnForum 		= (Request("forum") = "ON")
	
	IF NOT (strFirstName="" OR strLastName="" OR strEmail="" OR strPass="" OR strPass2="") Then 'check that all fields are complete
		IF (strPass = strPass2) Then							'check that password fields match
			blnOkay = True
		Else
			'password fields dont match
			strError = "pass"
		End If
	Else
		'all fields not entered
		strError = "fields"
	End If
	
	If blnOkay Then
		'add user to db
		strSQL = "EXECUTE CUK_UPDATEUSER @fn='"&strFirstName &"', @ln='"&strLastName &"', @e='" & strEmail &"', @p='"&strPass &"', @n="&Int(blnNews)&", @f="&Int(blnForum)&", @id=" & Session("ID")
		Set rs = cn.Execute(strSQL)
		Set rs = cn.Execute("EXECUTE CUK_GETUSER @un='"&Session("uName")&"'")
		intID = Int(rs("ID"))
		Set rs = Nothing
	
		Session("firstName") = strFirstName
		Session("lastName") = strLastName
		Session("uname") = strUname
		Session("email") = strEmail
		cn.close
		response.Redirect("/account/loginOut.asp?message=Details+sucessfully+changed")
	End If
End If

Set rs = cn.Execute("EXECUTE CUK_GETUSER @un='"&Session("uName")&"'")
strPassword = rs("password")
blnNews = (rs("news") = "True")
blnForum= (rs("forum")= "True")
Set rs = Nothing
%>
<H2>Edit your details</H2>
<P align="center">Here you can change your details that you originally entered.</P>
<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.name.value == "")
  {
    alert("Please enter a value for the \"First name\" field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.name.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"First name\" field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.name.value.length > 50)
  {
    alert("Please enter at most 50 characters in the \"First name\" field.");
    theForm.name.focus();
    return (false);
  }

  if (theForm.lname.value == "")
  {
    alert("Please enter a value for the \"Last name\" field.");
    theForm.lname.focus();
    return (false);
  }

  if (theForm.lname.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Last name\" field.");
    theForm.lname.focus();
    return (false);
  }

  if (theForm.lname.value.length > 50)
  {
    alert("Please enter at most 50 characters in the \"Last name\" field.");
    theForm.lname.focus();
    return (false);
  }

  if (theForm.email.value == "")
  {
    alert("Please enter a value for the \"Email address\" field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.email.value.length < 5)
  {
    alert("Please enter at least 5 characters in the \"Email address\" field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.email.value.length > 100)
  {
    alert("Please enter at most 100 characters in the \"Email address\" field.");
    theForm.email.focus();
    return (false);
  }

  if (theForm.pass.value == "")
  {
    alert("Please enter a value for the \"Password\" field.");
    theForm.pass.focus();
    return (false);
  }

  if (theForm.pass.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Password\" field.");
    theForm.pass.focus();
    return (false);
  }

  if (theForm.pass.value.length > 50)
  {
    alert("Please enter at most 50 characters in the \"Password\" field.");
    theForm.pass.focus();
    return (false);
  }

  if (theForm.pass2.value == "")
  {
    alert("Please enter a value for the \"Password (confirmation)\" field.");
    theForm.pass2.focus();
    return (false);
  }

  if (theForm.pass2.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Password (confirmation)\" field.");
    theForm.pass2.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><FORM action="editProfile.asp" METHOD="post" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
  <DIV align="center">
    <CENTER>
    <%
	IF strError <> "" Then
		SELECT Case strError
			Case "email"		Response.Write "<p><FONT color=red><i>Your email appears to not be valid. Please re-enter.</i></font></P>"
			Case "fields"		Response.Write "<p><FONT color=red><i>Please use ALL fields.</i></font></P>"
			Case "pass"			Response.Write "<p><FONT color=red><i>Your passwords appear not to match. Please re-enter.</i></font></P>"
			Case "uname"		Response.Write "<p><FONT color=red><i>Sorry. That username is already taken. Please select another.</i></font></P>"
			Case Else			Response.Write "<p><FONT color=red><i>Sorry. An unknown error occured whils processing your application. Please contact the webmaster.</i></font></P>"
	   End Select
	  End If
	%>
    <TABLE border="0" cellpadding="0" cellspacing="0" style="border:1px solid #612B83; border-collapse: collapse" bordercolor="#111111" id="AutoNumber1">
      <TR>
        <TD width="100%"><IMG border="0" src="../../images/main_menus/editdetails.gif"></TD>
      </TR>
      <TR>
        <TD width="100%" background="../../images/grad_write_purple_large.gif">
        <DIV align="center">
          <CENTER>
          <TABLE border="0" cellpadding="2" cellspacing="0" width="0" style="border-collapse: collapse" bordercolor="#111111">
            <TR>
              <TD>
              <P align="right">First name</P>
              </TD>
              <TD align="left">
              <P align="left">
              <!--webbot bot="Validation" s-display-name="First name" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="50" --><INPUT type="text" name="name" size="17" maxlength="50" value="<%=Session("firstName")%>"></P>
              </TD>
            </TR>
            <TR>
              <TD align="right">
              <P>Last name</P>
              </TD>
              <TD align="left">
              <P>
              <!--webbot bot="Validation" s-display-name="Last name" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="50" --><INPUT type="text" name="lname" size="17" maxlength="50" value="<%=Session("lastName")%>"></P>
              </TD>
            </TR>
            <TR>
              <TD align="right">
              <P>E-Mail address</P>
              </TD>
              <TD align="left">
              <P>
              <!--webbot bot="Validation" s-display-name="Email address" b-value-required="TRUE" i-minimum-length="5" i-maximum-length="100" --><INPUT type="text" name="email" size="17" maxlength="100" value="<%=Session("email")%>"></P>
              </TD>
            </TR>
            <TR>
              <TD align="right">
              <P>Username</P>
              </TD>
              <TD>
              <P><B><%=Session("uName")%></B></P>
              </TD>
            </TR>
            <TR>
              <TD align="right">
              <P>Password</P>
              </TD>
              <TD>
              <P align="center">
              <!--webbot bot="Validation" s-display-name="Password" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="50" --><INPUT type="password" name="pass" maxlength="50" size="17" style="float: left"></P>
              </TD>
            </TR>
            <TR>
              <TD align="right">
              <P>Password (confirm)</P>
              </TD>
              <TD>
              <P align="center">
              <!--webbot bot="Validation" s-display-name="Password (confirmation)" b-value-required="TRUE" i-minimum-length="1" --><INPUT type="password" name="pass2" size="17" style="float: left"></P>
              </TD>
            </TR>
            <TR>
              <TD align="right">
              <P>Subscribe to newsletter? </P>
              </TD>
              <TD>
              <P align="left"><INPUT type="checkbox" name="news" value="ON" <%if blnNews Then%>checked<%End If%>> </P>
              </TD>
            </TR>
            <TR>
              <TD align="right">Have my forum message<BR>
&nbsp;replies sent to me?</TD>
              <TD><INPUT type="checkbox" name="forum" value="ON" <%if blnForum Then%>checked<%End If%>></TD>
            </TR>
          </TABLE>
          </CENTER>
        </DIV>
        <P align="center"> 
    <INPUT type="submit" value="Edit details &gt; &gt;" name="B1" class="button" ><BR>
&nbsp;</P>
        </TD>
      </TR>
    </TABLE>
    </CENTER>
  </DIV>
  <INPUT type="hidden" name="uname" value="<%=Session("uName")%>">
  <INPUT type="hidden" name="submit_form" value="true">
</FORM>
<!--#include virtual="/includes/footer.asp" -->