<%
Option Explicit
strTitle="Register"
Dim blnJustForm
Dim cn, strFirstName, strLastName, strEmail, strUname, strPass, strPass2, strNews, blnNews, strExtra
Dim blnUserOk, blnOkay, intID, blnForum, strError
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" --><%
strFirstName 	= strIntoDB(Request("name"))
strLastName 	= strIntoDB(Request("lname"))
strEmail 		= strIntoDB(Request("email"))
strUname 		= Replace(strIntoDB(Request("uname")), "%", "")
strPass  		= Replace(strIntoDB(Request("pass")), "%", "")
strPass2 		= Replace(strIntoDB(Request("pass2")), "%", "")
blnNews  		= (Request("news")  = "ON")
blnForum 		= (Request("forum") = "ON")
	
blnOkay = False
blnUserOk = False

If Request("submit_form") = "true" Then
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open strDBMod

	'Validate Form.
	IF NOT (strFirstName="" OR strLastName="" OR strEmail="" OR strUname="" OR strPass="" OR strPass2="") Then 'check that all fields are complete
		'Check for duplicates
		Set rs = cn.Execute("EXECUTE CUK_GETUSER @un='"&strUname&"'")
		IF rs.EOF Then
			blnUserOk = True
		Else
			strError = "uname"
		End If
		Set rs = Nothing
		IF strPass = strPass2 Then	'check that password fields match
			blnOkay = True
		Else
			'password fields dont match
			strError = "pass"
		End If
	Else
		'all fields not entered
		strError = "fields"
	End If
	
	If blnOkay AND blnUserOk Then
		'add user to db
		strSQL = "EXECUTE CUK_REGISTER @fn='"&strFirstName&"', @ln='"&strLastName&"', @un='"&strUname&"', @p='"&strPass&"', @e='"&strEmail&"', @n="&Int(blnNews)&", @f="&Int(blnForum)
		Set rs = cn.Execute( strSQL )
		Set rs = cn.Execute("EXECUTE CUK_GETUSER @un='"&strUname&"'")
		intID = Int(rs("ID"))
		Set rs = nothing
		cn.Close
	
		Session("firstName") = strFirstName
		Session("lastName") = strLastName
		Session("uname") = strUname
		Session("name") = Session("firstName") & " " & Session("lastName")
		Session("email") = strEmail
		Session("logged") = True
		Session("ID") = intID
		Session("numLoggedIn") = 1
		Response.cookies("cocktailHeavenMembersUserName") = strUname
		Response.cookies("cocktailHeavenMembersUserName").Expires = "December 1, 2020"
		If Request("sendto") <> "" Then
			Response.Redirect(Request("sendto"))
		End If
	
		response.Redirect("/account/login.asp")
		Session("first") = ""
	End If

	cn.Close
	Set cn = Nothing
End If

blnJustForm = (Request("justform") = "true")
%>
<%if blnJustForm  Then%>
	<LINK href="/style/style.css" type="text/css" rel="stylesheet">
	<BODY class="nobackground">
<%Else%>
 <!--#include virtual="/includes/header.asp" -->
 <H2>Register with Cocktail : UK</H2>
 <table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
   <tr>
     <td width="100%"> <P style="align: justify" align="justify">When you are a member of Cocktail : UK you are able submit your own recipes and, more usefully,&nbsp; save your own personal ingredients list from which you can make/find cocktails from (you won&#39;t believe 
 the amount of recipes you can make from a small amount of ingredients!). </P>
<%End If%> </P>
 <P align="center">Please read our <A target="_blank" href="/services/privacy.asp">privacy policy</A> at your leisure. </P>
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

  if (theForm.uname.value == "")
  {
    alert("Please enter a value for the \"Username\" field.");
    theForm.uname.focus();
    return (false);
  }

  if (theForm.uname.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"Username\" field.");
    theForm.uname.focus();
    return (false);
  }

  if (theForm.uname.value.length > 50)
  {
    alert("Please enter at most 50 characters in the \"Username\" field.");
    theForm.uname.focus();
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
//--></script><!--webbot BOT="GeneratedScript" endspan --><FORM action="register.asp" METHOD="POST" target="_top" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
   <%If Request("sendto") <> "" Then%> <INPUT TYPE="HIDDEN" VALUE="<%=Request.ServerVariables("HTTP_REFERER")%>&amp;type=<%=Replace(Request("type"), "'", "")%>" name="sendto"><%End If%>
   <DIV align="center">
     <CENTER><%
	IF NOT strError = "" Then
		SELECT Case strError
			Case "email"		Response.Write "<p><FONT color=red><i>Your email appears to be invalid. Please re-enter.</i></font></P>"
			Case "fields"		Response.Write "<p><FONT color=red><i>Please use ALL fields.</i></font></P>"
			Case "pass"			Response.Write "<p><FONT color=red><i>Your passwords appear not to match. Please re-enter.</i></font></P>"
			Case "uname"		Response.Write "<p><FONT color=red><i>Sorry. That username is already taken. Please select another.</i></font></P>"
			Case Else			Response.Write "<p><FONT color=red><i>Sorry. An unknown error occured whilst processing your application. Please contact the webmaster.</i></font></P>"
	   End Select
	  End If
	%>
     <TABLE border="0" cellpadding="0" cellspacing="0" style="border:1px solid #612B83; border-collapse: collapse" bordercolor="#111111" id="AutoNumber1">
       <TR>
         <TD width="100%"><IMG border="0" src="../../images/main_menus/registeroncocktailuk.gif"></TD>
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
               <P align="left"><INPUT type="text" name="name" size="17" maxlength="50" value="<%=strFirstName%>"></P>
               </TD>
             </TR>
             <TR>
               <TD align="right">
               <P>Last name</P>
               </TD>
               <TD align="left">
               <P>
               </TD>
             </TR>
             <TR>
               <TD align="right">
               <P>E-Mail address</P>
               </TD>
               <TD align="left">
               <P><INPUT type="text" name="email" size="17" maxlength="100" value="<%=strEmail%>"></P>
               </TD>
             </TR>
             <TR>
               <TD align="right">
               <P>Username</P>
               </TD>
               <TD align="left">
               <P align="center"><INPUT name="uname" size="17" maxlength="50" value="<%=strUname%>" style="float: left"></P>
               </TD>
             </TR>
             <TR>
               <TD align="right">
               <P>Password</P>
               </TD>
               <TD align="left">
               <P align="center"><INPUT type="password" name="pass" maxlength="50" size="17" style="float: left"></P>
               </TD>
             </TR>
             <TR>
               <TD align="right">
               <P>Password (confirm)</P>
               </TD>
               <TD align="left">
               <P align="center">
               </TD>
             </TR>
             <TR>
               <TD align="right">
               <P>Subscribe to newsletter? </P>
               </TD>
               <TD>
               <P align="left"><INPUT type="checkbox" name="news" value="ON" checked> </P>
               </TD>
             </TR>
           </TABLE>
           </CENTER>
         </DIV>
         <P align="center">
    <INPUT type="submit" value="Register &gt; &gt;" name="B1" class="button" ><BR>
         <A href="login.asp<%If Request("sendto") <> "" Then%>?justform=true&sendto=<%=Request.ServerVariables("HTTP_REFERER")%><%End If%>">Already a member?</A><BR>
&nbsp;</P>
         </TD>
       </TR>
     </TABLE>
     </CENTER>
   </DIV>
   <INPUT type="hidden" name="submit_form" value="true">
 </FORM>
</td>
   </tr>
</table>
 <%if NOT blnJustForm  Then%>
 <!--#include virtual="/includes/footer.asp" -->
 <%End If%></BODY>