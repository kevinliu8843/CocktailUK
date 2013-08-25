<%
Sub sendCocktailsubmitEmail(strName, strEmail)
	Dim strBody, strSubject
	strSubject = "Cocktail:UK Drink Submission Confirmation"
	strBody= "<HTML><HEAD><TITLE>Cocktail:UK Drink Submission Confirmation</TITLE><LINK href=""http://www.cocktail.uk.com/style/mail.css"" type=""text/css"" rel=""stylesheet"" /></HEAD><BODY bgcolor=#ffffff><P>Hi,<BR>The cocktail you submitted, " & strName & ", has been viewed and has been added to the Cocktail : UK recipe list.<BR>Thank you for taking the time to submit the recipe, your input is highly appreciated and is what helps make the site.<BR><A href=""http://www.cocktail.uk.com"">http://www.cocktail.uk.com</A></BODY></HTML>"
	call SendEmail("theteam@cocktail.uk.com", strEmail, "", "", strSubject, strBody, True, "")
End Sub
%>