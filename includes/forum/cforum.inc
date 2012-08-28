<%
Dim strForumUsername, intForumStatus

Const ASCENDING				= 1
Const DESCENDING			= 2
Const SUBJECT_ORDER			= 4
Const POSTER_ORDER			= 8
Const DATE_ORDER			= 16
Const USERNAME_ORDER		= 32
Const FIRSTNAME_ORDER		= 64
Const LASTNAME_ORDER		= 128
Const EMAIL_ORDER			= 256
Const STATUS_ORDER			= 512
Const POSTS_ORDER			= 1024
Const VIEWED_ORDER			= 2048

Const FORUM_AS_TREE			= 1
Const FORUM_AS_LIST			= 2


' Cookie stuff
Const FORUM_COOKIE_FILENAME		= "COCKTAILUK_FORUMS" 
Const FORUM_MESSAGE_ORDERING	= "MESSAGE_ORDERING"
Const FORUM_MESSAGE_DISPLAY		= "MESSAGE_DISPLAY"
Const FORUM_THREAD_ORDERING		= "THREAD_ORDERING"


if Session("logged") then
	strForumUsername = Session("uname")
	if bIsAdmin Then
		intForumStatus = 1 ' We are a moderator...
	else
		intForumStatus = 0 ' We are NOT a moderator...
	End If
else
	strForumUsername  = ""
End if


Class CForum
	Public m_objConn				' The connection string

	' Member variables that are contained within the Forums table ----
	Public m_intID
	Public m_intNumThreads
	Public m_intNumMessages

	Public m_intStatus
	Public m_strName
	Private m_blnWAP
	'----------------------------------------------------------------
	Public m_bIsEditing				' True if we are in editing mode

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Public Sub Class_Initialize
		set m_objConn = Server.CreateObject("ADODB.Connection")
		m_objConn.Open(strDB)
		m_bIsEditing = False
		m_blnWAP = False
	End Sub

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Public Function GetRecordSet(rs)
		m_intID				= rs("ID")
		m_intNumThreads		= rs("NumThreads")
		m_intNumMessages	= rs("NumMessages")
		m_intStatus			= rs("Status")
		m_strName			= Trim(rs("Name"))
	End Function
	
'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Public Function ListAllForums(rs)
		Dim sql
		Dim bHighlight
		Dim strLine, i

		sql = "SELECT * FROM Forums ORDER BY ID"
		rs.Open sql, m_objConn, 0, 3
		
		ListAllForums = ListAllForums & "<table cellpadding=""0"" cellspacing=""1"" border=""0"">"
		
		strLine = "<tr><td colspan=""5""><hr size=""1"" noshade style=""color: Silver;""></td></tr>"
		ListAllForums = ListAllForums & strLine

		ListAllForums = ListAllForums & "<tr bgcolor=#ebe2f7><td>&nbsp;</td>"
		ListAllForums = ListAllForums & "<td>" & "Subjects&nbsp;&nbsp;&nbsp;</td>"
		ListAllForums = ListAllForums & "<td>" & "Messages&nbsp;&nbsp;&nbsp;</td>"
		ListAllForums = ListAllForums & "<td colspan=""2"">" & "Last&nbsp;Reply&nbsp;&nbsp;&nbsp;</td></tr>"
		
		ListAllForums = ListAllForums & strLine
		bHighlight = True
		i=0
		Do While NOT rs.EOF
			i = i + 1
			If i<3 Then
			GetRecordSet(rs)
			
			If bHighlight Then
				ListAllForums = ListAllForums & "<tr bgcolor=""#F0F0F0"">"
			Else
				ListAllForums = ListAllForums & "<tr>"
			End If

			ListAllForums = ListAllForums & "<td><b>" & SubjectHREF(m_intID, m_bIsEditing)
			ListAllForums = ListAllForums & m_strName & "</a>&nbsp;&nbsp;&nbsp;</b></td>"
			
			If m_intNumThreads >0 Then
				ListAllForums = ListAllForums & "<td>" & m_intNumThreads & "</td>"
			Else
				ListAllForums = ListAllForums & "<td>" & "-</td>"
			End If

			If m_intNumMessages >0 Then
				ListAllForums = ListAllForums & "<td>" & m_intNumMessages & "</td>"
			Else
				ListAllForums = ListAllForums & "<td>" & "-</td>"
			End If

			' Determine the last post ----------------------------------------------------
			sql = "SELECT TOP 1 DatePosted FROM ForumMessages "
			sql = sql & "WHERE ParentID IS NOT NULL And ForumID=" & m_intID & " ORDER BY DatePosted DESC"
			rs2.Open sql, m_objConn, 0, 3
			If Not rs2.EOF Then
				ListAllForums = ListAllForums & "<td>" & FormatDateTime(rs2("DatePosted"), 4) & "&nbsp;</td>"
				ListAllForums = ListAllForums & "<td>" & FriendlyDate(rs2("DatePosted")) & "</td>"
			Else
				ListAllForums = ListAllForums & "<td colspan=""3"">" & "-</td>"
			End If
			rs2.Close
			'----------------------------------------------------------------------------

			ListAllForums = ListAllForums & "</tr>"

			bHighlight = Not bHighlight
			End if
			rs.MoveNext
		Loop

		ListAllForums = ListAllForums & strLine
		ListAllForums = ListAllForums & "</table>"
		
		rs.Close
	End Function

	Public Function CompactID(strID)
		' Compacts GUID style IDs
		If Not IsNullID(strID) Then
			CompactID = Mid(strID, 2, 8)
			CompactID = CompactID & Mid(strID, 11, 4)
			CompactID = CompactID & Mid(strID, 16, 4)
			CompactID = CompactID & Mid(strID, 21, 4)
			CompactID = CompactID & Mid(strID, 26, 12)
		Else
			CompactID = "0"
		End If
	End Function

	Public Function ExpandID(strID)
		' Expands previously Compacted GUID style IDs
		If Len(strID)  = 32 Then
			ExpandID = "{" & Mid(strID, 1, 8)
			ExpandID = ExpandID & "-" & Mid(strID, 9, 4)
			ExpandID = ExpandID & "-" & Mid(strID, 13, 4)
			ExpandID = ExpandID & "-" & Mid(strID, 17, 4)
			ExpandID = ExpandID & "-" & Mid(strID, 21, 12) & "}"
		Else
			ExpandID = "{00000000-0000-0000-0000-000000000000}"
		End If
	End Function

	Public Function IsNullID(strID)
		IsNullID = (Len(strID) <> 38)
	End Function

	Public Function MessageLink(intForumID, strThreadID, strOrder, intPageNumber, strSearchList, bIsEditing)
		If bIsEditing = True Then
			MessageLink = "/admin/forums/edit_messages.asp"
		Else
			MessageLink = "/forums/messages.asp"
		End If

		MessageLink = MessageLink & "?fid=" & intForumID & "&tid=" & CompactID(strThreadID)
		
		If strOrder <> "" Then
			MessageLink = MessageLink & "&order=" & strOrder 
		End If

		If intPageNumber <> 0 Then
			MessageLink = MessageLink & "&pg=" & intPageNumber
		End If
		
		If strSearchList <> "" Then
			MessageLink = MessageLink & "&sl=" & Server.URLEncode(strSearchList)
		End If
		MessageLink = MessageLink & """"
	End Function

	Public Function MessageHREF(intForumID, strThreadID, strOrder, intPageNumber, strSearchList, bIsEditing)
		MessageHREF = "<a href=""" & MessageLink(intForumID, strThreadID, strOrder, intPageNumber, strSearchList, bIsEditing) & ">"
	End Function

	Public Function ThreadLink(intForumID, strOrder, intPageNumber, bIsEditing)
		If bIsEditing = True Then
			ThreadLink = "/admin/forums/edit_subjects.asp"
		Else
			ThreadLink = "/forums/subjects.asp"
		End If
		
		ThreadLink = ThreadLink & "?fid=" & intForumID
		
		If strOrder <> "" Then
			ThreadLink = ThreadLink & "&order=" & strOrder 
		End If

		If intPageNumber <> 0 Then
			ThreadLink = ThreadLink & "&pg=" & intPageNumber
		End If

		ThreadLink = ThreadLink & """"
	End Function

	Public Function ThreadHREF(intForumID, strOrder, intPageNumber, bIsEditing)
		ThreadHREF = "<a href=""" & ThreadLink(intForumID, strOrder, intPageNumber, bIsEditing) & ">"
	End Function

	Public Function SearchLink(intForumID, strOrder, intPageNumber, strQuery, strOption)
		SearchLink = "/forums/search.asp"
		SearchLink = SearchLink & "?fid=" & intForumID
		
		If strOrder <> "" Then
			SearchLink = SearchLink & "&order=" & strOrder 
		End If

		If intPageNumber <> 0 Then
			SearchLink = SearchLink & "&pg=" & intPageNumber
		End If

		SearchLink = SearchLink & "&o=" & strOption
		SearchLink = SearchLink & "&q=" & Server.URLEncode(strQuery)

		SearchLink = SearchLink & """"
	End Function

	Public Function SearchHREF(intForumID, strOrder, intPageNumber, strQuery, strOption)
		SearchHREF = "<a href=""" & SearchLink(intForumID, strOrder, intPageNumber, strQuery, strOption) & ">"
	End Function

	Public Function SubjectLink(intForumID, bIsEditing)
		If bIsEditing = True Then
			SubjectLink = "/admin/forums/edit_subjects.asp?fid=" & intForumID & """"
		Else
			SubjectLink = "/forums/subjects.asp?fid=" & intForumID & """"
		End If
	End Function

	Public Function SubjectHREF(intForumID, bIsEditing)
		SubjectHREF = "<a href=""" & SubjectLink(intForumID, bIsEditing) & ">"
	End Function

	Public Function ReplyLink(intForumID, strThreadID, bIsEditing)
		If bIsEditing = True Then
			ReplyLink = "/admin/forums/edit_message.asp?fid=" & intForumID & "&mid=" & CompactID(strThreadID) & """"
		Else
			ReplyLink = "/forums/reply.asp?fid=" & intForumID & "&mid=" & CompactID(strThreadID) & """"
		End If
	End Function

	Public Function ReplyHREF(intForumID, strThreadID, bIsEditing)
		ReplyHREF = "<a href=""" & ReplyLink(intForumID, strThreadID, bIsEditing) & ">"
	End Function

	Public Function FormatURL(strURL)
		' Strips out bits from our URL that we don't want...
		Dim intStart, intEnd

		If Instr( strURL, "forums/configure.asp") > 0 Then
			FormatURL = "/forums/"
		Else 
			intStart = Instr(strURL, "order")

			If intStart > 0 Then
				FormatURL = Left(strURL, intStart-1)

				intEnd = Instr(intStart, strURL, "&")
				If intEnd > 0 Then
					FormatURL = FormatURL & Mid(strURL, intEnd+1)
				End If
			Else
				FormatURL = strURL
			End If
		End If

		If Trim(FormatURL) = "" Then
			FormatURL = "/forums/"
		End If
	End Function

	Private Function FilterSwearWord( strBody, strSwearWord , bRemoveWord )
		' If bRemoveWord  = true	then kill swear word
		' If bRemoveWord  = false	then replace with *'s

		Dim strLowerCaseBody, intPos, intStartPos
		Dim strLeftChar, strRightChar

		strLowerCaseBody	= LCase( strBody )
		strSwearWord		= LCase( strSwearWord )
		FilterSwearWord		= strBody

		intStartPos = 1
		Do
			intPos = InStr(intStartPos, strLowerCaseBody, strSwearWord )

			If intPos <> 0 Then 
				strLeftChar		= " "
				strRightChar	= " "

				If intPos > 1 Then
					strLeftChar = LCase( Mid( FilterSwearWord, intPos-1, 1 ) )
				End If
				
				If intPos + Len( strSwearWord ) <= Len( FilterSwearWord ) Then
					strRightChar = LCase( Mid( FilterSwearWord, intPos+Len(strSwearWord), 1 ) )
				End If

				If (strLeftChar >= "a" AND strLeftChar <= "z" ) OR (strRightChar >= "a" AND strRightChar <= "z" ) Then
					' This might not be a swear word.
					intStartPos = intPos + Len( strSwearWord )
				Else
					If bRemoveWord Then
						FilterSwearWord		= Left( FilterSwearWord, intPos-1 ) & Mid( FilterSwearWord, intPos + Len(strSwearWord) )
						strLowerCaseBody	= Left( strLowerCaseBody, intPos-1 ) & Mid( strLowerCaseBody, intPos + Len(strSwearWord) )
						intStartPos			= intPos
					Else
						FilterSwearWord		= Left( FilterSwearWord, intPos-1 ) & String(Len(strSwearWord), "*") & Mid( FilterSwearWord, intPos + Len(strSwearWord) )
						strLowerCaseBody	= Left( strLowerCaseBody, intPos-1 ) & String(Len(strSwearWord), "*") & Mid( strLowerCaseBody, intPos + Len(strSwearWord) )
						intStartPos			= intPos
					End If
					
				End If
			End If		
		Loop While intPos <> 0
	End Function

	'**********************************************************************************************************
	'**********************************************************************************************************
	'**********************************************************************************************************

	Public Function FilterALLSwearWords( strBody, bRemoveWord )
		FilterALLSwearWords = strBody

		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "shit", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "fuck", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "fucking", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "fucked", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "piss", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "pissed", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "cunt", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "fucker", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "motherfucker", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "cocksucker", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "nigger", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "anus", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "arse", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "ass", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "muthafucker", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "mutha-fucker", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "bitch", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "slut", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "bollocks", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "whore", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "gimp", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "bugger", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "penis", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "tits", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "muff", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "todger", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "lesbo", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "wanker", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "wank", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "blow job", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "clit", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "ballbags", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "cum ", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "cumming", bRemoveWord )
		FilterALLSwearWords  = FilterSwearWord( FilterALLSwearWords, "gaylord", bRemoveWord )
	End Function

	Public Sub setWAP()
		m_blnWAP = True
	End Sub
	
	Public Function StripHTMLSpecial(strBody)
		Dim strTAG, strHTML
		Dim aryDontFilter(2)
		Dim iStart, iEnd, iPos
		Dim bKeep
		If NOT m_blnWAP Then
			aryDontFilter(0) 	= "<br>"	'Set array limitlessly with tags we don't want filtered
			aryDontFilter(1) 	= "<p>"
		Else
			strBody = Replace(strBody, "<br>", "<br/>", 1, -1, 1)
			aryDontFilter(0) 	= "<br/>"
		End If
		bKeep				= False
		
		StripHTMLSpecial = strBody
		iStart 	= Instr(1, StripHTMLSpecial, "<")
		iEnd	= Instr(1, StripHTMLSpecial, ">")
		Do While iStart <> 0 And iEnd <> 0
			If iStart > iEnd Then 					'If '>' is found before '<' the string is corrupt.  '>' will be ignored
				iStart = iEnd + 1 					'adjust cursor to after '>' tag
			Else
				strTAG = Mid(StripHTMLSpecial, iStart, (iEnd-iStart+1)) 'Retrieving HTML tag
				For Each strHTML In aryDontFilter 	'comparing tag to tags we don't want filtered
					If strTAG = strHTML Then
						bKeep = True
					End If
				Next
				If bKeep = True Then
					iStart				= iEnd + 1 	'If we want to keep tag, just change string cursor
					bKeep 				= False 
				Else
					StripHTMLSpecial = Left(StripHTMLSpecial, iStart-1) & Right(StripHTMLSpecial, Len(StripHTMLSpecial) - iEnd) 'remove tag
				End If
			End If
			iPos = iStart							'setting string cursor
			If iPos = 0 Then iPos = 1
			iStart 	= Instr(iPos, StripHTMLSpecial, "<")
			iEnd	= Instr(iPos, StripHTMLSpecial, ">")
		Loop
	End Function

	Public Function FriendlyDate(strDate)
		' Changes, say, Sunday, 1st July 2000 into Sunday 1st July

		FriendlyDate = WeekdayName(Weekday(strDate), True)
		FriendlyDate = FriendlyDate & " " & MonthName(Month(strDate), True)
		FriendlyDate = FriendlyDate & " " & Day(strDate)
		If Year(strDate) <> Year(Now()) Then
			FriendlyDate = FriendlyDate & " " & Year(strDate)
		End If
	End Function

End Class
%>