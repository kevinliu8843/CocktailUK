<%
Function displayRating( strRating, intUsers )
	'Split up RatingsAndUsers
	Dim numRating, strPerson, intFraction, i, j

	numRating = strRating
	intUsers = Int(intUsers)
	If intUsers = 1 THEN
		strPerson = "person has "
	ELSE
		strPerson = "people have "
	End If
	intFraction = Right( CStr( numRating * 10), 1 )
	intFraction = intFraction*2

	Response.write("<TABLE BORDER=""0"" CELLPADDING=""0"" CELLSPACING=""0""><TR><TD nowrap>")
	Response.Write("" & intUsers & " " & strPerson & "rated this: &nbsp;</TD><TD>" & VbCrLf)

	Call displayRatingGraphOnly( strRating )

	Response.write("</TD><TD nowrap> &nbsp;&nbsp;<B>" & numRating & "</B> out of 5</TD></TR></TABLE>")
End Function

Sub displayRatingGraphOnly( strRating )
	'Split up RatingsAndUsers
	Dim intFraction, i, j, blnTable, imgsWritten, partValue
	intFraction = Right( CStr( Int(strRating) * 10), 1 )
	intFraction = intFraction*2
	blnTable = False

	If blnTable Then
		Response.Write("<TABLE BORDER=""2"" CELLPADDING=""0"" CELLSPACING=""0""><TR HEIGHT=""4"">" & VbCrLf)
	
		For i=1 to Int(strRating)
			Response.Write("<TD bgcolor=""lightgrey""><IMG SRC=""/images/redpix.gif"" WIDTH=""20"" HEIGHT=""4"" /></TD>" & VbCrLf)
		Next
		
		IF NOT Int(strRating) = 5 AND NOT intFraction = 0 THEN
			Response.Write("<TD bgcolor=""lightgrey""><IMG SRC=""/images/redpix.gif"" WIDTH=""" & intFraction & """ HEIGHT=""4""><IMG SRC=""/images/pixel.gif"" WIDTH=""" & 20 - intFraction & """ HEIGHT=""4""></TD>" & VbCrLf)
		END IF
	
		If NOT Int(strRating) = 5 AND intFraction = 0 THEN
			Response.Write("<TD bgcolor=""lightgrey""><IMG SRC=""/images/pixel.gif"" WIDTH=""" & 20 - intFraction & """ HEIGHT=""4""></TD>" & VbCrLf)
		END IF
	
		IF NOT Int(strRating) = 5 THEN
			For j=1 to ( 4-Int(Int(strRating)) )
				Response.Write("<TD bgcolor=""lightgrey""><IMG SRC=""/images/pixel.gif"" WIDTH=""20"" HEIGHT=""4"" /></TD>" & VbCrLf)
			Next
		END IF
		
		Response.write("</TR></TABLE>")
	Else
		imgsWritten = 0
		partValue = CDbl( CDbl(strRating) - Fix(CDbl(strRating)) )

		For i=1 to Fix(CDbl(strRating))
			Response.Write("<IMG src=""/images/sitesearch/1.gif"" height=12 width=13 alt=""Rated: "&strRating&""" border=""0"">")
			imgsWritten = imgsWritten + 1
		Next

		If (partValue) > 0.21 AND (partValue) < 0.79 Then
			Response.Write("<IMG src=""/images/sitesearch/0.5.gif"" height=12 width=13 alt=""Rated: "&strRating&""" border=""0"">")
			imgsWritten = imgsWritten + 1
		End If

		If (partValue) >= 0.79 Then
			Response.Write("<IMG src=""/images/sitesearch/1.gif"" height=12 width=13 alt=""Rated: "&strRating&""" border=""0"">")
			imgsWritten = imgsWritten + 1
		End If
		
		For i=imgsWritten to 5 - 1
			Response.Write("<IMG src=""/images/sitesearch/0.gif"" height=12 width=13 alt=""Rated: "&strRating&""" border=""0"">")
		Next
	End If
End Sub

Function addRating( strCurrentUsers, strCurrentRate, strRateAdded,  strID, objConn )
	'calculate rating
	Dim numRating, rate

	numRating = ( ( Int(strCurrentUsers) * CDbl(strCurrentRate) ) + ( CDbl(strRateAdded) ) ) / ( Int(strCurrentUsers) +1 )
	strSQL = "UPDATE cocktail SET users='"&  Int( strCurrentUsers ) + 1 & "', rate='" & Round( numRating, 1 ) & "' WHERE ID=" & strID
	Set rs = cn.Execute( strSQL )

	'close objects
	rs2.Close
	Set rs2 = Nothing
	Set rs = Nothing
	cn.Close
	Set cn = Nothing
	Response.Redirect("/cocktails/recipe.asp?rate=true&ID=" & strID)
End Function
%>
