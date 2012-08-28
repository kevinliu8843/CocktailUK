<%
Class CForumAutoReply
	Public m_objConn

	Public m_objForum
	Public m_intForumID
	Public m_strThreadID
	Public m_strMessageID
	Public m_intUID

	Public Sub Class_Initialize
		Set m_objForum	= New CForum
		Set m_objConn	= Server.CreateObject("ADODB.Connection")
		m_objConn.Open(strDBMod)
	End Sub

	Public Sub Class_Terminate
		Set m_objForum	= Nothing
		Set m_objConn	= Nothing
	End Sub

	Public Sub AddToAutoReply()
		'Not used at all in Cocktail : UK site
		Dim sql

		sql = "INSERT INTO ForumAutoReply (FID, ThreadID, UID) VALUES ("
		sql = sql & m_intForumID & ", "
		sql = sql & "'" & m_objForum.ExpandID(m_strThreadID) & "', "
		sql = sql & m_intUID
		sql = sql & ")"

		'm_objConn.Execute sql
	End Sub

	Public Sub SendAutoReplies(strSubject)
		'Don't include m_intUID in this email (as they are the one that just submitted it)
		Dim sql, strBody, strUserNames 
		Dim strFirstname, strLastname, strEmail
		Dim strWholeName, strParentID
		Dim objEmail

		sql = "SELECT Members.Firstname, Members.Lastname, Members.email FROM Members "
		sql = sql & "INNER JOIN ForumAutoReply ON ForumAutoReply.UID = Members.UID "
		sql = sql & "WHERE ForumAutoReply.FID = " & m_intForumID
		sql = sql & " AND ForumAutoReply.ThreadID = '" & m_objForum.ExpandID(m_strThreadID) & "'"
		sql = sql & " AND ForumAutoReply.UID <> " & m_intUID
		
		sql = "SELECT DISTINCT username, ParentID FROM ForumMessages "
		sql = sql & "WHERE ForumMessages.ForumID = " & m_intForumID
		sql = sql & " AND ForumMessages.ThreadID = '" & m_objForum.ExpandID(m_strThreadID) & "'"
		
		rs.Open sql, m_objConn, 0, 3
		strUserNames = ""
		Do While Not rs.EOF
			strParentID = rs("ParentID")
			If Session("uname") <> rs("username") then
				strUserNames = strUserNames & "'" & rs("username") & "',"
			End If
			rs.MoveNext
		Loop
		strUserNames = strUserNames & "'" 
		rs.Close
		
		sql = "SELECT firstName, lastName, email, userName, forum from usr where userName IN ("&strUserNames&"') AND forum=1"

		rs.Open sql, m_objConn, 0, 3

		Do While Not rs.EOF
			strFirstname	= Trim(rs("firstName"))
			strLastname		= Trim(rs("lastName"))
			strEmail		= Trim(rs("email"))
			
			If strEmail <> "" Then
				If strFirstname <> "" Then
					strWholeName = strFirstname & " " & strLastname
				Else
					strWholeName = "Registered User"
				End If
	
				strBody = "Dear " & strWholeName & "," & VbCrLf & VbCrLf
				strBody = strBody & "A reply has been posted to the following Cocktail : UK forum subject:-" & VbCrLf & VbCrLf
				strBody = strBody & """" & strSubject & """" & VbCrLf & VbCrLf
	
				strBody = strBody & "To view the reply go to the following page:-" & VbCrLf
	
				strBody = strBody & "http://www.cocktail.uk.com/forums/messages.asp?fid=" & m_intForumID & "&tid=" & m_strMessageID & VbCrLf & VbCrLf
	
				strBody = strBody & "Thank you for participating in our forums." & VbCrLf & VbCrLf
	
				strBody = strBody & "Should you not wish to recieve further notifications," & VbCrLf
				strBody = strBody & "please login to your member's area and edit your profile." & VbCrLf

				strBody = strBody & "The Cocktail : UK Team" & VbCrLf
				strBody = strBody & "http://www.cocktail.uk.com" & VbCrLf
	
				' ---------------------------------------------------------
				call SendEmail("forums@cocktail.uk.com", strEmail, "", "", "Cocktail : UK Forum Update", strBody, False, "")
				' ---------------------------------------------------------
			End If
			rs.MoveNext
		Loop
		rs.Close

	End Sub
End Class
%>