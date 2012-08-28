<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/CAPTCHA/CAPTCHA_process_form.asp" -->
<%
strDomain = "http://www.cocktail.uk.com/"

Select Case Request("type")
	Case "cocktail"
		'Get form elements
		strFromName = Request("name1")
		strFromEmail = Request("email1")
		strFriendName = Request("name2")
		strFriendEmail = Request("email2")
		strComments = Request("comments")
		
		strCocktailName = Request("cocktailName")
		strDirections = Request("directions")
		strIngredients = Request("ingredients")
		strServes = Request("serves")
		strCategory = Request("category")
		strImgsrc = Request("imgsrc")
		
		set cn = Server.CreateObject("ADODB.Connection")
		cn.open strDB
		cn.execute("EXECUTE CUK_ADDCOUNTEREVENT @col='sentcocktails'")
		cn.close
		set cn = nothing
		
		id = Request("id")
		
		blnConfirm  = (Request("confirm") = "ON")		
		blnHTML = (Request("html") = "ON")
		blnHTML = True 'Force this
		
		strBodyText = "<HTML><HEAD><TITLE>An e-cocktail for you from Cocktail : UK</TITLE><LINK rel=""stylesheet"" type=""text/css"" href="""&strDomain&"style/style.css""></HEAD>"
		strBodyText = strBodyText &"<BODY> <DIV align=""center"">  <CENTER>   <TABLE border=""0"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse: collapse"" bordercolor=""#111111"" id=""AutoNumber1"">"
		strBodyText = strBodyText &"<TR>       <TD valign=""top"">       <H3 align=""center"">"&strCocktailName&"<BR>       <A href=""http://www.cocktail.uk.com/db/viewcocktail.asp?ID="&id&""">"&strImgsrc&"</A></H3>"
		strBodyText =  strBodyText &"      </TD>       <TD valign=""top"" width=""400"">       <H2 align=""center""><I>You have received an <BR>       e-cocktail from Cocktail : UK</I></H2>"
		strBodyText = strBodyText  &"      <P>Dear "&strFriendName&",<BR>       <A href=""mailto:"&strFromEmail&""">"&strFromName&"</A> has sent you a cocktail from Cocktail : UK called &quot;"&strCocktailName&"&quot;.</P>       <P>Why don't you pop along to <A href=""http://www.cocktail.uk.com/db/viewCocktail.asp?ID="&id&""">http://www.cocktail.uk.com</A> "
		strBodyText = strBodyText &"to view the recipe for yourself...</TD>     </TR>   </TABLE>   </CENTER> </DIV></BODY></HTML>"
		
		Response.cookies("CocktailUKSendDrinkName1") = strFromName
		Response.cookies("CocktailUKSendDrinkName1").Expires = "December 1, 2010"
		Response.cookies("CocktailUKSendDrinkName2") = strFriendName 
		Response.cookies("CocktailUKSendDrinkName2").Expires = "December 1, 2010"
		Response.cookies("CocktailUKSendDrinkEmail1") = strFromEmail 
		Response.cookies("CocktailUKSendDrinkEmail1").Expires = "December 1, 2010"
		Response.cookies("CocktailUKSendDrinkEmail2") = strFriendEmail
		Response.cookies("CocktailUKSendDrinkEmail2").Expires = "December 1, 2010"
		
		strSubject = strFromName & " has sent you an E-Cocktail from Cocktail : UK."
		If blnHTML Then
			strBody = strBodyText
		Else
			strBody = "Dear "&strFriendName&","&VbCrLf&strFromName&" has sent you an E-Cocktail from www.cocktail.uk.com for the pleasure of your palette and eyes!. "&vbCrLf&"You can view it at http://www.cocktail.uk.com/db/viewCocktail.asp?ID="&id&"."&VbCrLf&"There are a whole host of features on the site and most cocktail recipes come complete with an image."&VbCrLf&" You can also register and manage your own bar and add cocktails to the database."&VbCRLf&VbCRLf&"Many thanks"&VbCRLf&"The Team @ Cocktail : UK"
		End If
		
'Disabled due to spammers...
		'call SendEmail("theteam@cocktail.uk.com", strFriendEmail, "", "", strSubject, strBody, blnHTML, "")
	
		'If blnConfirm Then
		'	strSubject = "E-Mail confirmation from www.cocktail.uk.com"
		'	strBody = "Dear "&strFromName&","&VbCrLf&"An e-cocktail has been sent to "&strFriendName&"."&VbCrLf&VbCrLf&"Thank you for using this service."&VbCrLf&"Lee"&VbCrLf&"www.cocktail.uk.com"
		'	call SendEmail("theteam@cocktail.uk.com", strFromEmail, "", "", strSubject, strBody, False, "")
		'End If
		
	Case "friend"
		'Get form elements
		strFromName = Request("name")
		strFromEmail = Request("replyemail")
		strFriendName = Request("friendname")
		strFriendEmail = Request("friendemail")
		strComments = Request("comments")
	
		blnConfirm = (Request("confirm") = "ON")		
		blnHTML = (Request("html") = "ON")
		
		strSubject = strFromName & " recommends this site."
		If blnHTML Then
			strBody = "<HTML><HEAD><TITLE>"&strFromName&" recommends this site.</TITLE><LINK href="""&strDomain&"style/mail.css"" type=""text/css"" rel=""stylesheet"" /></HEAD><BODY><P>Dear "&strFriendName&",<BR>"&strFromName&" ("&strFromEmail&") has recommended <A href=""http://www.cocktail.uk.com"">www.cocktail.uk.com</A> to you. Why not take a look.<P>There are a whole host of features on the site. Most cocktail recipes come complete with an image. You can also register and manage your own bar and add cocktails to the database."
			If strComments & "" <> "" Then
				strBody = strBody & "<P><B>Their comments : </B>" & strComments & VbCrLf & VbCrLf
			End If
			strBody = strBody & "<P>Many thanks<BR>The team at <A href=""http://www.cocktail.uk.com"">www.cocktail.uk.com</A></BODY></HTML>"
		Else
			strBody = "Dear "&strFriendName&","&VbCrLf&strFromName&" has recommended www.cocktail.uk.com to you. Why not take a look."&VbCrLf&"There are a whole host of features on the site. Most cocktail recipes come complete with an image. You can also register and manage your own bar and add cocktails to the database."&VbCRLf&VbCRLf
			If strComments & "" <> "" Then
				strBody = strBody & "Their comments : " & strComments & VbCrLf & VbCrLf
			End If
			strBody = strBody & "Many thanks"&VbCRLf&"www.cocktail.uk.com"
		End If
'Disabled due to spammers...
		'call SendEmail(strFromEmail, strFriendEmail, "", "", strSubject, strBody, blnHTML, "")

		'If blnConfirm Then
		'	strSubject = "E-Mail confirmation from www.cocktail.uk.com"
		'	strBody = "Dear "&strFromName&","&VbCrLf&"An e-mail has been sent to "&strFriendName&" recommending www.cocktail.uk.com."&VbCrLf&VbCrLf&"Thank you for using this service. It is much appreciated."&VbCrLf&"www.cocktail.uk.com"
		'	call SendEmail("theteam@cocktail.uk.com", strFromEmail, "", "", strSubject, strBody, False, "")
		'End If
		
	Case "details"	
		'Get user's details:
		Set cn = Server.CreateObject("ADODB.Connection")
		cn.Open strDB
		strSQL = "SELECT * FROM usr WHERE email='" & Trim( Replace(Request("email"), "'", "") ) & "'"
		Set rs = cn.Execute( strSQL )
		IF NOT rs.EOF Then
			strUsername = rs("userName")
			strPassword = rs("password")
			strFirstName = rs("firstName")
		Else
			response.Redirect ("/mail/forgottenPassword.asp?error=email")
		End If
		rs.Close
		cn.Close
		Set cn = Nothing
		Set rs = Nothing
		
		strMailTo 	= Request("email")
		strMailFrom	= "theteam@cocktail.uk.com"
		strSubject 	= "Cocktail : UK username and password"
		strBody 	= "<HTML><HEAD><TITLE>Cocktail.uk.com Username and Password</TITLE><LINK href="""&strDomain&"style/mail.css"" type=""text/css"" rel=""stylesheet"" /></HEAD><BODY><H2>Your Cocktail.uk.com Username and Password</H2>  <P>Dear "&strFirstName&",</P><P>Your username and password for <A href=""http://www.cocktail.uk.com"">www.cocktail.uk.com</A> are as follows:</P><P><B>Username : </B>" & strUsername & "<BR/><B>Password : </B>" & strPassword & "</P><P>Type them in as shown above as the login script is case sensitive.</P><P><A href=""http://www.cocktail.uk.com"">www.cocktail.uk.com</A></P></BODY></HTML>"
	
		call SendEmail(strMailFrom, strMailTo, "", "", strSubject, strBody, True, "")

	Case "page"
		'Get form elements
		strFromName = Request("name")
		strFromEmail = Request("replyemail")
		strFriendName = Request("friendname")
		strFriendEmail = Request("friendemail")
		strComments = Request("comments")
		strUrl = Request("url")
	
		strMailTo = strFriendEmail
		Mail.From = "theteam@cocktail.uk.com"
		Mail.To = strMailTo
		Mail.Subject = strFromName & " thinks you may find this page from Cocktail : UK interesting..."
		strBody = "Dear "&strFriendName&","&VbCrLf&strFromName&" ("&strFromEmail&") thinks you may find this page ("&strUrl&") from Cocktail : UK (www.cocktail.uk.com) interesting... "&vbCrLf
		strBody = strBody & "Why not take a look."&VbCrLf
		If strComments <> "" Then
			strBody = strBody & "Here is what "& strFromName &" said about the page: """ & strComments & """"&VbCrLf
		End If
		strBody = strBody & "There are a whole host of features on the site. Most cocktail recipes come complete with an image. You can also register and manage your own bar and add cocktails to the database."&VbCRLf&VbCRLf&"Many thanks"&VbCRLf&"Cocktail : UK (www.cocktail.uk.com)"
	
	Case Else
		If NOT blnCAPTCHAcodeCorrect Then
			If Request("from") <> "" Then
				Response.redirect(Request("from") & "?fail=Please enter the correct code in the box provided. Your email has NOT been sent to us.")
			Else
				Response.redirect("/services/contact.asp?fail=Please enter the correct code in the box provided. Your email has NOT been sent to us.")
			End If
		End If

		strBody = "<HTML><HEAD><TITLE>cocktail.uk.com suggestion</TITLE><LINK href="""&strDomain&"style/mail.css"" type=""text/css"" rel=""stylesheet"" /></HEAD><BODY><P><B>"&Request("name")&"</B> (<a href=""mailto:"&Request("replyemail")&""">"&Request("replyemail")&"</A>) has made a suggestion : <P>"&Replace(Request("comments"), VbCrLf, "<BR>" & VbCrLf)&"</BODY></HTML>"
		If Request("mailto") <> "" Then
			strMailTo = Request("mailto") & "@cocktail.uk.com"
		else
			strMailTo = "theteam@cocktail.uk.com"
		end if
		strSubject = "Cocktail.uk.com suggestion"
		call SendEmail(Request("replyemail"), strMailTo, "", "", strSubject, strBody, True, "")
End Select

if Request("type") = "cocktail" then
	response.redirect("/db/viewCocktail.asp?ID="&id&"&mail=true")
elseif Request("type") <> "" Then
	Response.Redirect ("/mail/thankYou.asp?type="&Request("type"))
else
	Response.Redirect ("/mail/thankYou.asp")
end if
%>