<%
Option Explicit
strTitle="Your Account"
Dim cn, strConfirm, blnAccepted, blnCookie, blnCookieDeleted, setupCookie, blnJustform, cookie
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%

response.buffer = True

IF request.queryString("logout") = "True" Then
	Session.Contents.RemoveAll
	Response.Redirect("/")
End If

'Test for submission
IF NOT Request("uname") = "" AND NOT Request("pass") = "" Then
	Set cn = Server.CreateObject("ADODB.Connection")
	cn.Open strDBMod
	strSQL = "EXECUTE CUK_LOGIN @strUser='" & strIntoDB(Replace(Trim( Request("uname") ),"%","" )) & "', @strPassword='" & strIntoDB(Replace(Trim( Request("pass") ),"%","" )) & "'"
	Set rs = cn.Execute( strSQL )
	IF NOT rs.EOF Then
		blnAccepted = True
		
		Session("firstName") = strOutDB(rs("firstName"))
		Session("lastName") = strOutDB(rs("lastName"))
		Session("name") = Session("firstName") & " " & Session("lastName") 
		Session("uname") = strOutDB(rs("userName"))
		Session("email") = strOutDB(rs("email"))
		Session("password") = Replace(Trim( rs("password") ),"%","" )
		Session("logged") = True
		Session("ID") = rs("ID")
		Session("numLoggedIn") = rs("loggedIn")
		
		Response.cookies("user") = Session("email")
		Response.cookies("user").Expires = "December 1, 2010"
		
		If Request("sendto") <> "" Then
			rs.Close
			cn.Close
			Set cn = Nothing
			Set rs = Nothing
			Response.Redirect(Request("sendto"))
		End If
		
		IF NOT Request("ID") = "" Then
			rs.Close
			cn.Close
			Set cn = Nothing
			Set rs = Nothing
			response.Redirect("/cocktails/recipe.asp?ID="&Request("ID"))
		End If
	Else
		strConfirm = "<FONT color=red>Username or password incorrect. Please re-enter your details</font>"
		blnAccepted = False
	End If
	rs.Close
	cn.Close
	Set cn = Nothing
	Set rs = Nothing
Else
	blnAccepted = False
End If
%>
<!--#include virtual="/includes/header.asp" -->
<%IF NOT Session("logged") AND NOT blnAccepted Then%>
    <H2>Login</H2>
    <FORM method="POST" action="login.asp" name="form1" target="_top">
    <%IF NOT Request.Querystring("ID") = "" Then%>
        <INPUT TYPE=HIDDEN VALUE=<%=Request.QueryString("ID")%> NAME=ID>
    <%End If%>
    <%If Request("sendto") <> "" Then%>
        <INPUT TYPE=HIDDEN VALUE="<%=Replace(Request("sendto"), "'", "")%>&type=<%=Replace(Request("type"), "'", "")%>&mid=<%=Request("mid")%>" name=sendto>
    <%End If%>
    <%If NOT blnJustForm Then%>
    <%If Request("doublecheck") = "true" Then%>
        <P align="center"><FONT color=red><I>As a part of our updated security, pelase re-enter your password. You won't need to do this again.<BR>
        To cancel your auto-login, please <A href="login.asp?clearcookie=true">click here</A>.</I></FONT></P>
        <INPUT TYPE=HIDDEN VALUE="true" name="dontlogin">
    <%End If%>
    <p align="center">In order to access Cocktail : UK members features, you need to login.<P>
    <%End If%>
    <DIV align="center">
      <%=strConfirm%>
      <CENTER>
      <TABLE border="0" cellpadding="0" cellspacing="0" style="border:1px solid #612B83; border-collapse: collapse" bordercolor="#612B83" id="AutoNumber1">
        <TR>
          <TD align="center">
    <IMG border="0" src="../../images/main_menus/logintoyouraccount.gif"></TD>
        </TR>
        <TR>
          <TD align="center" background="../../images/grad_write_purple.gif">
    <TABLE border="0" cellpadding="5" cellspacing="0" style="text-align: center" height="121">
      <TR>
        <TD height="23"><B>Username:</B></TD>
        <TD height="23">
            <INPUT type="text" name="uname" size="20" value="<%=Request("uname")%>">
        </TD>
      </TR>
      <TR>
        <TD height="23"><B>Password:</B></TD>
        <TD height="23">
            <INPUT type="password" name="pass" size="20">
            </TD>
      </TR>
      <TR>
        <TD colspan="2" height="16"> 
        <INPUT type="submit" value="Login &gt; &gt;" name="B1" class="button" ></TD>
      </TR>
      <TR>
        <TD colspan="2" height="19"><A target="_top" href="/mail/password.asp">Forgotten your username or password?</A></TD>
      </TR>
    </TABLE>
          </TD>
        </TR>
      </TABLE>
      </CENTER>
    <p>If you do not have an account, please click <A target="_top" href="/account/register.asp">HERE</A><br>
    If you are trying to check the status of your shop order, please click
    <a target="_top" href="http://www.drinkstuff.com/member/secure/login.asp">HERE</a></DIV>
    </FORM>
    <SCRIPT language="javascript">
	    document.form1.uname.focus()
    </script>

<%
Else
	'do cookie stuff
	Response.cookies("user") = Session("email")
	Response.cookies("user").Expires = Now()+365*10

	If (Int( Session("numLoggedIn") ) = 1) AND (Session("first") = "") Then
		Response.Redirect("/account/firstVisit.asp")
	End If
%>
<h2> <%=Session("firstname")%>'s Account</h2>
<p><%=strConfirm%>
<%
	If Request("message") <> "" Then
%>
<P align="center"><FONT color="#FF0000"><I><%=Request("message")%></I></FONT></P>
<%
	End If
%>
<P>Hi <%=Session("firstname")%>,<BR>
Welcome to your section of the site. Here you can store you own personal
ingredients that match your home bar in your very own web-bar. From this you can
find out what you can make from our list of drinks. But as you can see, there is
more to your members section that this, so go explore!!!</p>

<P align="center">
<map name="FPMap10">
<area href="selectIngredients.asp" shape="rect" coords="16, 48, 164, 77">
<area href="userHotList.asp" shape="rect" coords="219, 66, 370, 96">
<area href="whatCanIMake.asp" shape="rect" coords="40, 104, 238, 133">
<area href="/cocktails/random.asp" shape="rect" coords="84, 144, 399, 181">
<area href="../randomGenerator.asp" shape="rect" coords="22, 190, 272, 222">
<area href="submitCocktail.asp" shape="rect" coords="218, 226, 393, 258">
<area href="editProfile.asp" shape="rect" coords="8, 246, 159, 279">
<area href="/shop/products/affiliate.asp" coords="82, 287, 390, 322" shape="rect"></map>
<img border="0" src="../../images/members_area.gif" usemap="#FPMap10" width="400" height="326">

<P align="center"><A href="?logout=True">Logout</A><BR>&nbsp;</P>
<%
End If
%><!--#include virtual="/includes/footer.asp" -->
