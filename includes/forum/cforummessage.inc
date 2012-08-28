<%
Class CForumMessage
	Public m_objConn, m_objConnUpdate
	Private m_objForum

	' Member variables that are contained within the ForumMessages table ----
	Public m_strID				' ID of this Message
	Public m_strParentID		' ID of the direct parent.  Can be NULL
	Public m_strThreadID		' ID of the first ancestor.  i.e Parent of Parent...
	Public m_intUID				' ID of the user who submitted this message
	Public m_intForumID			' 
	Public m_intNumChildren		' Number of direct children this message has
	Public m_intNumMessages		' Number of total messages this message has beneath it (children, grandchildren, etc.)
	Public m_intStoryID			' Does this message belong to a story item?

	Public m_strSearchList		' List of comma-delimited words that we have searched for (so we can highlight)

	Public m_intStatus			' 0 = Delete, 1 = Normal
	Public m_datePosted			'
	Public m_strSubject			'
	Public m_strUsername		' Copy of Username from Members database (for quickness)
	Public m_strMessage

	'----------------------------------------------------------------
	Public m_intOrdering
	Public m_intDisplayStyle		' FORUM_AS_TREE / FORUM_AS_LIST

	Public m_bHighlight				' Internal 'global' variable, used to toggle displaying of background colour or not

	' Paging variables
	Public m_intPageSize			' The number of messages to display per page
	Public m_intCurrentPage			'

	Public m_bIsEditing				' True if we are in editing mode

	Public Sub Class_Initialize
		Set m_objForum	= New CForum
		Set m_objConn	= Server.CreateObject("ADODB.Connection")
		m_objConn.Open(strDB)
		Set m_objConnUpdate	= Server.CreateObject("ADODB.Connection")
		m_objConnUpdate.Open(strDBMod)

		' Clear off database fields ------------------
		m_strID				= ""
		m_strParentID		= ""
		m_strThreadID		= ""
		m_intUID			= 0
		m_intForumID		= 0
		m_intNumChildren	= 0
		m_intNumMessages	= 0
		m_intStoryID		= 0
		m_intStatus			= 1
		m_datePosted		= Now()
		m_strSubject		= ""
		m_strUsername		= ""
		m_strMessage		= ""
		'---------------------------------------------

		m_intOrdering		= DATE_ORDER Or DESCENDING		' This will be stored in the cookie
		m_intDisplayStyle	= FORUM_AS_TREE

		m_strSearchList		= ""

		' Paging variables
		m_intPageSize				= 6
		m_intCurrentPage			= 1

		m_bIsEditing				= False
	End Sub

	Public Sub Class_Terminate
		Set m_objForum	= Nothing
		Set m_objConn	= Nothing
		Set m_objConnUpdate = Nothing
	End Sub

	Public Function GetRecordSet(rs)
		GetRecordSet = True
		m_strID				= rs("ID")
		m_strParentID		= rs("ParentID")
		m_strThreadID		= rs("ThreadID")
		'm_intUID			= rs("UID")
		m_intForumID		= rs("ForumID")
		m_intNumChildren	= rs("NumChildren")
		m_intNumMessages	= rs("NumMessages")
		m_intStoryID		= rs("StoryID")
		m_intStatus			= rs("Status")
		m_datePosted		= rs("DatePosted")
		m_strSubject		= rs("Subject")
		m_strSubject		= HighlightAllWords(m_strSubject)
		m_strUsername		= rs("Username")
		m_strMessage		= rs("Message")
		m_strMessage		= HighlightAllWords(m_strMessage)
	End Function

	Sub GetDetailsFromMessageID()
		' Given an ID get the ForumID and ThreadID

		Dim sql
		If m_strID <> "" Then
			sql = "SELECT TOP 1 ThreadID, ForumID, subject From ForumMessages WHERE ID = '" & m_strID & "' AND Status = 1"
			rs.Open sql, m_objConn, 0, 3
			If Not rs.EOF Then
				m_strThreadID	= rs("ThreadID")
				m_intForumID	= rs("ForumID")
				m_strSubject	= rs("Subject")
				m_strSubject	= HighlightAllWords(m_strSubject)
			End If
			rs.Close
		ElseIf m_strThreadID <> "" Then
			sql = "SELECT TOP 1 subject From ForumMessages WHERE ThreadID = '" & m_strThreadID & "' AND Status = 1"
			rs.Open sql, m_objConn, 0, 3
			If Not rs.EOF Then
				m_strSubject	= rs("Subject")
				m_strSubject	= HighlightAllWords(m_strSubject)
			End If
			rs.Close
		End If
	End Sub

	Private Function HighlightAllWords(strText)
		Dim aryWords, i
		If m_strSearchList <> "" Then
			aryWords = Split(m_strSearchList, ",")
			i=0
			Do While i<=UBound(aryWords)
				HighlightAllWords	= HighlightWord(strText, aryWords(i))
				i=i+1
			Loop
		Else
			HighlightAllWords = strText
		End If
		HighlightAllWords = hyperlink(HighlightAllWords)
	End Function

	Private Function HighlightWord(strText, strSearchWord)
		Dim strLowerCase, intFoundPos, intStart, intLenSearchWord
		Dim i
		Dim strHTMLBefore, strHTMLAfter

		strHTMLBefore		= "<span style=""background-color: #0000FF;color: #FFFFFF;"">"
		strHTMLAfter		= "</span>"

		intLenSearchWord	= Len(strSearchWord)
		strSearchWord		= LCase(strSearchWord)
		strLowerCase		= LCase(strText)
		intStart			= 1
		HighlightWord		= strText
		For i=1 to 10
			intFoundPos	= Instr(intStart, strLowerCase, strSearchWord)
			If intFoundPos = 0 Then Exit For
			
			HighlightWord	= Left(HighlightWord, intFoundPos-1) & strHTMLBefore & Mid(HighlightWord,intFoundPos, intLenSearchWord) & strHTMLAfter & Right(HighlightWord, Len(HighlightWord)-(intFoundPos+(intLenSearchWord-1)))
			strLowerCase	= Left(strLowerCase, intFoundPos-1) & strHTMLBefore & Mid(strLowerCase,intFoundPos, intLenSearchWord) & strHTMLAfter & Right(strLowerCase, Len(strLowerCase)-(intFoundPos+(intLenSearchWord-1)))

			intStart = intFoundPos + intLenSearchWord + Len(strHTMLBefore & strHTMLAfter)
		Next
	End Function

	Public Sub SaveSettings()
		Response.Cookies(FORUM_COOKIE_FILENAME)(FORUM_MESSAGE_ORDERING)	= m_intOrdering
		Response.Cookies(FORUM_COOKIE_FILENAME)(FORUM_MESSAGE_DISPLAY)	= m_intDisplayStyle
	End Sub

	Public Sub LoadSettings()
		If Request.Cookies(FORUM_COOKIE_FILENAME).HasKeys Then
			If Request.Cookies(FORUM_COOKIE_FILENAME)(FORUM_MESSAGE_ORDERING) <> "" Then
				m_intOrdering		= Request.Cookies(FORUM_COOKIE_FILENAME)(FORUM_MESSAGE_ORDERING)
			End If
			If Request.Cookies(FORUM_COOKIE_FILENAME)(FORUM_MESSAGE_DISPLAY) <> "" Then
				m_intDisplayStyle	= Request.Cookies(FORUM_COOKIE_FILENAME)(FORUM_MESSAGE_DISPLAY)
			End If
		End If
	End Sub
	
	' Entry point routine for recursive (tree-view) of sub-messages
	Public Function DisplayAsTree(rs)
		Dim sql, threadID
		sql = ""
		' Display the initial message (top-level message)
		If m_strThreadID <> "" Then
			sql = "SELECT ID, ParentID, ThreadID, ForumID, NumChildren, NumMessages, StoryID, Status, datePosted, Subject, Username, Message From ForumMessages WHERE ForumID = " & m_intForumID & " And ID = '" & m_strThreadID & "' AND Status = 1"
		ElseIf m_intStoryID >0 Then
			sql = "SELECT ID, ParentID, ThreadID, ForumID, NumChildren, NumMessages, StoryID, Status, datePosted, Subject, Username, Message From ForumMessages WHERE ForumID = " & m_intForumID & " And StoryID = '" & m_intStoryID & "' AND Status = 1"
		End If

		If sql <> "" Then
			rs.Open sql, m_objConn, 0, 3
			If Not rs.EOF Then
				threadID= rs("ID")
				DisplayAsTree = True
				GetRecordSet(rs)
				rs.Close

				If m_bIsEditing = True Then
					%>
					<form action="<%=m_objForum.MessageLink(m_intForumID, m_strThreadID, GetOrder(), m_intCurrentPage, "", m_bIsEditing)%>" method="post">
					<input type="submit" name="action" value="Delete">
					<%
				End If

				Call LinkToOriginalItem()
				Call DisplaySingleMessageAsTree(True, 0)

				m_bHighlight = True
				Call ACTUAL_DisplayAsTree(rs, m_strThreadID, 1)

				If m_bIsEditing = True Then
					%></form><%
				End If
			Else
				DisplayAsTree = False
				rs.Close
			End If
		Else
			DisplayAsTree = False
		End If
		If threadID <> ""  Then
			m_objConnUpdate.execute("UPDATE ForumMessages SET viewed=viewed+1 WHERE ThreadID='"&threadID&"'")
		End If
	End Function

	Public Sub LinkToOriginalItem()
'		Dim strSite
'		strSite = ""
'		If m_intEventID > 0 And (m_intForumID=1 Or m_intForumID=2) Then
'			Select Case m_intForumID
'				Case 1	strSite = g_strSiteDirectory(SITE_CLASSICALPLUS)
'				Case 2	strSite = g_strSiteDirectory(SITE_JAZZPLUS)
'			End Select
'
'			If strSite <> "" Then
'				Response.Write("<font color=""#840000"">&#149;</font>&nbsp;View details of this <a href=""" & strSite & "/webcasts/webcast.asp?id=" & m_intEventID & """>webcast</a><br>")
'			End If
'		ElseIf m_intNewsID > 0 Then
'			strSite = g_strSiteDirectory(g_intSite)
'			Response.Write("<font color=""#840000"">&#149;</font>&nbsp;View the original <a href=""" & strSite & "/news/default.asp?ms=" & ForumIDToMusicStyle(m_intForumID) & "&id=" & m_intNewsID & """>news</a> item<br>")
'
'		ElseIf m_intStoryID > 0 Then
'			strSite = g_strSiteDirectory(g_intSite)
'			Response.Write("<font color=""#840000"">&#149;</font>&nbsp;View the original <a href=""" & strSite & "/story.asp?sc=" & CONST_CHANNEL_FORUMS & "&id=" & m_intStoryID & """>story</a><br>")
'		End If
	End Sub

	Private Sub DisplaySingleMessageAsTree(bAllowReply, intDepth)
		' bAllowReply - Not sure if we're going to need this...
		' intDepth - If we're viewing as a tree this will tab the tree out
		' bHighlight - If True we display a background colour for this message
		' m_bHighlight = True
		%>
		<table <% If m_bHighlight Then%>bgcolor="#e0e0e0"<%End If%> cellpadding="0" cellspacing="0" border="0" width="100%">
		<tr>
		<% If intDepth > 0 Then %>
			<td rowspan="2" width="1%"><img src="/images/pixel.gif" width="<%=intDepth * 20%>" height="1"></td>
		<% End If %>

		<td rowspan="2" width="80%">
		<%If m_bIsEditing = True And (m_strID <> m_strThreadID) Then %>
			<input type="checkbox" name="ID" value="<%=m_objForum.CompactID(m_strID)%>">
		<%End If %>

		<%
		'If bAllowReply = True Then Response.Write(m_objForum.ReplyHREF(m_intForumID, m_strID, m_bIsEditing))
		Response.Write( "<B>" & m_strSubject & "</B>" )
		'If bAllowReply = True Then Response.Write("</A> " & m_objForum.ReplyHREF(m_intForumID, m_strID, m_bIsEditing) & "(Reply)</a>")
		%>
		</td>
		<% If Int(m_intDisplayStyle) = FORUM_AS_TREE Then %>
		<td width="20%" align="right">
		<%Else%>
		<td width="20%" align="left">
		<%End If%>
		
			<%=m_strUsername%>
		</td>
		</tr>
		<tr>
			<td width="20%" align="right">
				<%=m_objForum.FriendlyDate(m_datePosted) & " " & FormatDateTime(m_datePosted, 4)%>
			</td>
		</tr>
		<tr>
		<%If intDepth > 0 Then %>
			<td width="1%"><img src="/images/pixel.gif" width="<%=intDepth * 20%>" height="1"></td>
		<%End If%>

		<%If intDepth > 0 Then %>
		<td colspan="2">
		<%Else %>
		<td colspan="3">
		<%End If %>
			<%= m_strMessage%>
		</td>
		</tr>
		<tr>
			<td colspan="3">&nbsp;</td>
		</tr>
		</table>
<%
	End Sub

	Private Sub ACTUAL_DisplayAsTree(rs, strParentID, intDepth)
		Dim sql, bRecordsToDisplay, aryMessages, i, intNumMessages, bHaveMessages, intNumChildren

		If strParentID <> "" Then
		
			sql = "SELECT ID, ParentID, ThreadID, ForumID, NumChildren, NumMessages, StoryID, Status, datePosted, Subject, Username, Message "
			sql = sql & "FROM ForumMessages WHERE Status=1 And ForumID = " & m_intForumID & " AND ParentID = '" & strParentID & "' ORDER BY DatePosted"
			rs.Open sql, m_objConn, 0, 3
			
			If Not rs.EOF Then
				aryMessages = rs.GetRows
				bHaveMessages = True
			Else
				bHaveMessages = False
			End If
			rs.Close
				
			If bHaveMessages = True Then
				intNumMessages = UBound(aryMessages, 2)
				' For loops are shit in VBScript...
				i=0
				Do While i <= intNumMessages
					m_strID				= aryMessages(0, i)
					m_strParentID		= aryMessages(1, i)
					m_strThreadID		= aryMessages(2, i)
					m_intForumID		= aryMessages(3, i)
					m_intNumChildren	= aryMessages(4, i)
					m_intNumMessages	= aryMessages(5, i)
					m_intStoryID		= aryMessages(6, i)
					m_intStatus			= aryMessages(7, i)
					m_datePosted		= aryMessages(8, i)
					m_strSubject		= aryMessages(9,i)
					m_strSubject		= HighlightAllWords(m_strSubject)
					m_strUsername		= aryMessages(10,i)
					m_strMessage		= aryMessages(11,i)
					m_strMessage		= HighlightAllWords(m_strMessage)
	
					Call DisplaySingleMessageAsTree(True, intDepth)
					m_bHighlight = Not m_bHighlight		' Toggle background colour on/off
	
					If m_intNumChildren > 0 Then
						' We have some children, let's recurse here...
						Call ACTUAL_DisplayAsTree(rs, m_strID, intDepth + 1)
					End If
					i=i+1
				Loop
			End If
		End If
	End Sub

	' Displays all messages for a thread in a list, ordered by either subject, poster or date
	Public Function DisplayAsList(rs)
		Dim sql, i, strPageControls, intTotalPages
		
		' Create SQL ----------------------------------------------------------------------------
		' Display the initial message (top-level message)
		If m_strThreadID <> "" Then
			sql = "SELECT * From ForumMessages WHERE Status=1 And ForumID = " & m_intForumID & " AND ThreadID = '" & m_strThreadID & "' "
		ElseIf m_intStoryID >0 Then
			sql = "SELECT * From ForumMessages WHERE Status=1 And ForumID = " & m_intForumID & " AND StoryID = " & m_intStoryID
		End If

		If m_intOrdering And SUBJECT_ORDER Then
			sql = sql & "ORDER BY Subject"
		ElseIf m_intOrdering And POSTER_ORDER Then
			sql = sql & "ORDER BY Username"
		Else
			sql = sql & "ORDER BY DatePosted"
		End If
		
		If m_intOrdering And DESCENDING Then
			sql = sql & " DESC"
		End If


		' Handle pages ------------------------------------------------------------------------
		rs.PageSize			= m_intPageSize

		rs.CursorLocation	= 2
		rs.CacheSize		= m_intPageSize * 10

		rs.Open sql, m_objConn, 3, 3
		If Not rs.EOF Then
			rs.AbsolutePage = m_intCurrentPage
		End If
		
		intTotalPages	= rs.PageCount
		If intTotalPages > 1 Then
			strPageControls = GetPageControls(intTotalPages)
			Response.Write("<center>" & strPageControls & "<br></center>")
		End If

		' Display a page full of messages --------------------------------------------------------

		If m_bIsEditing = True Then
		%>
		<form action="<%=m_objForum.MessageLink(m_intForumID, m_strThreadID, GetOrder(), m_intCurrentPage, "", m_bIsEditing)%>" method="post">
		<input type="submit" name="action" value="Delete">
		<%
		End If

		Call DisplayHeader()
		%>
		<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#e0e0e0">
		<%
		DisplayAsList = False
		For i=1 to m_intPageSize
			If RS.EOF Then
				Exit For
			End If
			DisplayAsList = True

			GetRecordSet(rs)
			If m_bIsEditing = True Then
				' We need to have links on each message to edit them
				Call DisplaySingleMessageAsList(True, m_bHighlight)
			Else
				' We only need one "Reply To Subject" link to reply to this subject
				Call DisplaySingleMessageAsList(False, m_bHighlight)
			End If
			
			'm_bHighlight = Not m_bHighlight		' Toggle background colour on/off
			rs.MoveNext
		Next
		'----------------------------------------------------------------------------------------

		Response.Write("</table>")

		If m_bIsEditing = True Then
			Response.Write("</form>")
		End If
	End Function

	Public Sub SetOrder(strOrder)
		' strOrder will be a 2 byte string 
		' (1st character = direction)
		' (2nd character = sort type, i.e.)
		If Len(strOrder) <> 2 Then
			Exit Sub
		End If

		Select Case Mid(strOrder, 1, 1)
			Case "A"		m_intOrdering = ASCENDING
			Case Else		m_intOrdering = DESCENDING
		End Select

		Select Case Mid(strOrder, 2, 1)
			Case "S"		m_intOrdering = (m_intOrdering Or SUBJECT_ORDER)
			Case "P"		m_intOrdering = (m_intOrdering Or POSTER_ORDER)
			Case Else		m_intOrdering = (m_intOrdering Or DATE_ORDER)
		End Select
	End Sub

	Public Function GetOrder
		' Returns a string representing the current sort order (in the same format as SetOrder...)

		If (m_intOrdering And ASCENDING) Then
			GetOrder = "A"
		Else
			GetOrder = "D"
		End If

		If (m_intOrdering And SUBJECT_ORDER) Then
			GetOrder = GetOrder & "S"
		ElseIf (m_intOrdering And POSTER_ORDER) Then
			GetOrder = GetOrder & "P"
		Else
			GetOrder = GetOrder & "D"
		End If
	End Function

	Public Function AddMessage(rs)
		Dim sql, strThreadID

		If m_objForum.IsNullID(m_strThreadID) Then
			' We need to find the ThreadID that this child message belongs to
			sql = "SELECT TOP 1 ThreadID FROM ForumMessages WHERE ID='" & m_strParentID & "'"
			
			rs.Open sql, m_objConn, 0, 3
			If NOT rs.EOF Then
				m_strThreadID = rs("ThreadID")
				rs.Close
			Else 
				AddMessage = False
				rs.Close
				Exit Function
			End If
		End If

		sql = "INSERT INTO ForumMessages (ParentID, ThreadID, ForumID, NumChildren, NumMessages, StoryID, Status, ReIndex, DatePosted, Subject, Username, Message)"
		sql = sql & "VALUES ("

		If Not m_objForum.IsNullID(m_strParentID) Then					'CLSIDs must be 38 characters
			sql = sql & "'" & m_strParentID & "',"
		Else
			sql = sql & "NULL,"
		End If

		If Not m_objForum.IsNullID(m_strThreadID) Then					'CLSIDs must be 38 characters
			sql = sql & "'" & m_strThreadID & "',"
		Else
			sql = sql & "NULL,"
		End If

		m_strSubject	= m_objForum.StripHTMLSpecial(m_strSubject)
		m_strSubject	= m_objForum.FilterALLSwearWords( m_strSubject, False )
		m_strSubject	= Left(m_strSubject, 64)
		m_strSubject	= Replace(m_strSubject, "'", "''")

		If Not m_bIsEditing Then
			m_strMessage	= m_objForum.StripHTMLSpecial(m_strMessage)
		End If

		m_strMessage	= m_objForum.FilterALLSwearWords( m_strMessage, False )
		m_strMessage	= Replace(m_strMessage, "'", "''")
		m_strMessage	= Replace(m_strMessage, VBCrLf, "<br>")

		m_datePosted = Day(m_datePosted) & "/" & MonthName(Month(m_datePosted)) & "/" & Year(m_datePosted) & " " & Time()

		If m_strSubject <> "" And m_strMessage <> "" Then
		
			'sql = sql & m_intUID & ","
			sql = sql & m_intForumID		& ","
			sql = sql & m_intNumChildren	& ","
			sql = sql & m_intNumMessages	& ","
			sql = sql & m_intStoryID		& ","
			sql = sql & "1, "  ' Status
			sql = sql & "1, "  ' ReIndex
			sql = sql & "'" & m_datePosted	& "',"
			sql = sql & "'" & m_strSubject	& "',"
			sql = sql & "'" & m_strUsername	& "',"
			sql = sql & "'" & m_strMessage  & "'"
			sql = sql & ")"

			m_objConnUpdate.Execute(sql)
			
			' Update the 'latest post date' for this Thread

			' ----------------------------------------------------------------------------------------------------
			' This is a new child entry, we need to adjust certain fields in other records...
			' Add 1 to the number of children that this child-message's parent has now got
			'  - Also Update the latest time that this Thread was updated.
			sql = "UPDATE ForumMessages SET NumChildren = NumChildren + 1 WHERE ID='" & m_strParentID & "'"
			m_objConnUpdate.Execute(sql)
			
			' Add 1 to the number of messages that this child-message's thread has now got
			sql = "UPDATE ForumMessages SET NumMessages = NumMessages + 1, ThreadUpdated = '" & m_datePosted & "' WHERE ID='" & m_strThreadID & "'"
			m_objConnUpdate.Execute(sql)

			' Add 1 to the number of messages that this forum has now got
			sql = "UPDATE Forums SET NumMessages = NumMessages + 1 WHERE ID=" & m_intForumID
			m_objConnUpdate.Execute(sql)

			AddMessage = True
		Else
			AddMessage = False
		End If
	End Function

	Public Function UpdateMessage
		Dim sql

		m_strSubject	= m_objForum.StripHTMLSpecial(m_strSubject)
		m_strSubject	= Replace(m_strSubject, "'", "''")

		If Not m_bIsEditing Then
			m_strMessage	= m_objForum.StripHTMLSpecial(m_strMessage)
		End If
		m_strMessage	= Replace(m_strMessage, "'", "''")
		m_strMessage	= Replace(m_strMessage, VBCrLf, "<br>")
		
		sql = "UPDATE ForumMessages SET ReIndex=1, "
		sql = sql & "Subject = '" & m_strSubject & "',"
		sql = sql & "Message = '" & m_strMessage & "' "
		sql = sql & "WHERE ForumID = " & m_intForumID & " AND ID = '" & m_objForum.ExpandID(m_strID) & "'"

		m_objConnUpdate.Execute(sql)
	End Function

	Private Function GetPageControls(intTotalPages)
		Dim i

		If Int(m_intCurrentPage) > 1 Then
			GetPageControls = GetPageControls & m_objForum.MessageHREF(m_intForumID, m_strThreadID, GetOrder(), Int(m_intCurrentPage)-1, m_strSearchList, m_bIsEditing) & "&lt;&lt; Previous Page</a>&nbsp;"
		Else
			GetPageControls = GetPageControls & "&lt;&lt; Previous Page&nbsp;"
		End If

		For i=1 to intTotalPages
			If Int(i) = Int(m_intCurrentPage) Then
				GetPageControls = GetPageControls & "<b>" & i & "</b>&nbsp;"
			Else
				GetPageControls = GetPageControls & m_objForum.MessageHREF(m_intForumID, m_strThreadID, GetOrder(), i, m_strSearchList, m_bIsEditing) & i & "</a>&nbsp;"
			End If
		Next

		If Int(m_intCurrentPage) < Int(intTotalPages) Then
			GetPageControls = GetPageControls & m_objForum.MessageHREF(m_intForumID, m_strThreadID, GetOrder(), Int(m_intCurrentPage)+1, m_strSearchList, m_bIsEditing) & "Next Page &gt;&gt;</a>"
		Else
			GetPageControls = GetPageControls & "Next Page &gt;&gt;"
		End If
	End Function

	Private Sub DisplaySingleMessageAsList(bAllowReply, bHighlight)
		' bAllowReply - Not sure if we're going to need this...
		' bHighlight - If True we display a background colour for this message
%>
		<tr <% If bHighlight Then%>bgcolor="#F0F0F0"<%End If%>>
		<td rowspan="2" width="80%">
		<%If m_bIsEditing = True And (m_strID <> m_strThreadID) Then %>
			<input type="checkbox" name="ID" value="<%=m_objForum.CompactID(m_strID)%>">
		<%End If %>

		<%
		If bAllowReply = True Then
			Response.Write( m_objForum.ReplyHREF(m_intForumID, m_strID, m_bIsEditing))
		End If
		Response.Write( m_strSubject )

		If bAllowReply = True Then 
			Response.Write("</a>")
		End If
		%>

		</td>
		<% If Int(m_intDisplayStyle) = FORUM_AS_TREE Then %>
		<td width="20%" align="right">
		<%Else%>
		<td width="20%" align="left">
		<%End If%>
		
			<%= m_strUsername%>
		</td>
		</tr>
		<tr <% If bHighlight Then%>bgcolor="#F0F0F0"<%End If%>>
			<td width="20%" align="right">
				<%= m_objForum.FriendlyDate(m_datePosted) & ", " & Hour(m_datePosted) & ":" & Minute(m_datePosted) %>
			</td>
		</tr>

		<tr <% If bHighlight Then%>bgcolor="#F0F0F0"<%End If%>>
		<td colspan="3">
			<%= m_strMessage%>
		</td>
		</tr>
		<tr <% If bHighlight Then%>bgcolor="#F0F0F0"<%End If%>>
			<td colspan="3">
				&nbsp;
			</td>
		</tr>
<%
	End Sub

	Private Sub DisplayHeader()
		' Tidy this up, there's probably a nice, clean way of doing this!!!
		Dim strUpImg, strDownImg
		strUpImg	= "<img src=""/images/forums/up.gif"" width=""11"" height=""11"" border=""0"">"
		strDownImg	= "<img src=""/images/forums/down.gif"" width=""11"" height=""11"" border=""0"">"

		Call LinkToOriginalItem()
%>

		<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ebeedd">
		<tr>
			<td colspan="3"><hr size="1" noshade style="color: Silver;"></td>
		</tr>
		<tr>
		<td width="80%">
		<% 
		If (m_intOrdering AND SUBJECT_ORDER) Then 
			Response.Write("<b>Subject</b>") 
		Else
			Response.Write(m_objForum.MessageHREF(m_intForumID, m_strThreadID, "AS", m_intCurrentPage, m_strSearchList, m_bIsEditing) & "Subject</a>")
		End If

		If (m_intOrdering AND SUBJECT_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.MessageHREF(m_intForumID, m_strThreadID, "AS", m_intCurrentPage, m_strSearchList, m_bIsEditing) & strUpImg & "</a>")
			Else
				Response.Write(m_objForum.MessageHREF(m_intForumID, m_strThreadID, "DS", m_intCurrentPage, m_strSearchList, m_bIsEditing) & strDownImg & "</a>")
			End If
		End If
		
		If m_bIsEditing = False Then
			' When Editing is True we need to have links on each message to be able to edit them
			Response.Write( "&nbsp;(" & m_objForum.ReplyHREF(m_intForumID, m_strThreadID, m_bIsEditing) & "Reply</a>)<br>")
		End If

		%>
		</td>		
		<td width="10%">
		<% 
		If (m_intOrdering AND POSTER_ORDER) Then 
			Response.Write("<b>Poster</b>") 
		Else
			Response.Write(m_objForum.MessageHREF(m_intForumID, m_strThreadID, "AP", m_intCurrentPage, m_strSearchList, m_bIsEditing) & "Poster</a>")
		End If

		If (m_intOrdering AND POSTER_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.MessageHREF(m_intForumID, m_strThreadID, "AP", m_intCurrentPage, m_strSearchList, m_bIsEditing) & strUpImg & "</a>")
			Else
				Response.Write(m_objForum.MessageHREF(m_intForumID, m_strThreadID, "DP", m_intCurrentPage, m_strSearchList, m_bIsEditing) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>		

		<td width="10%" align="right">
		<% 
		If (m_intOrdering AND DATE_ORDER) Then 
			Response.Write("<b>Date</b>") 
		Else
			Response.Write(m_objForum.MessageHREF(m_intForumID, m_strThreadID, "AD", m_intCurrentPage, m_strSearchList, m_bIsEditing) & "Date</a>")
		End If

		If (m_intOrdering AND DATE_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.MessageHREF(m_intForumID, m_strThreadID, "AD", m_intCurrentPage, m_strSearchList, m_bIsEditing) & strUpImg & "</a>")
			Else
				Response.Write(m_objForum.MessageHREF(m_intForumID, m_strThreadID, "DD", m_intCurrentPage, m_strSearchList, m_bIsEditing) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>		
		</tr>
		<tr>
			<td colspan="3"><hr size="1" noshade style="color: Silver;"></td>
		</tr>
		</table>
<%
	End Sub

	Public Sub DeleteMessages(rs)
		Dim strID, sql, strParentID, strThreadID, intCount

		For Each strID in Request("ID")
			'Find the parent and thread of this soon to be deleted message
			sql = "SELECT TOP 1 ParentID, ThreadID FROM ForumMessages WHERE ID='" & m_objForum.ExpandID(strID) & "'"
			rs.Open sql, m_objConn, 0, 3
			If Not rs.EOF Then
				strParentID = rs("ParentID")
				strThreadID = rs("ThreadID")
				rs.Close

				' First Count the number of messages that will be "moved"
				sql = "SELECT COUNT(*) AS TheCount FROM ForumMessages WHERE ParentID = '" & m_objForum.ExpandID(strID) & "'"
				
				rs.Open sql, m_objConn, 0, 3
				If Not rs.EOF Then
					intCount = rs("TheCount")
				End If
				rs.Close
			
				' Update all direct children of this soon to be deleted message to have the parent of the deleted message
				sql = "UPDATE ForumMessages SET ParentID = '" & strParentID & "' WHERE ParentID = '" & m_objForum.ExpandID(strID) & "'"
				m_objConnUpdate.Execute sql

				' Update the number of child messages that the "new" parent has (-1 for the deleted message)
				sql = "UPDATE ForumMessages SET NumChildren = NumChildren + " & (intCount-1) & " WHERE ID = '" & strParentID & "'"
				m_objConnUpdate.Execute sql

				' Actually "delete" our message
				sql = "UPDATE ForumMessages SET Status = 0, ReIndex=0 WHERE ID = '" & m_objForum.ExpandID(strID) & "'"
				m_objConnUpdate.Execute sql
				
				' Update the number of total messages that this thread now has
				sql = "UPDATE ForumMessages SET NumMessages = NumMessages - 1 WHERE ID = '" & strThreadID & "'"
				m_objConnUpdate.Execute sql

				' Update the number of total messages that the Forum now has
				sql = "UPDATE Forums SET NumMessages = NumMessages - 1 WHERE ID =" & m_intForumID
				m_objConnUpdate.Execute sql
			Else
				rs.Close
			End If
			
		Next
	End Sub

	Public Function GetMessage(rs)
		Dim sql
		GetMessage = False
		sql=""
		If m_strID <> "" Then
			sql = "SELECT TOP 1 * From ForumMessages WHERE ID = '" & m_objForum.ExpandID(m_strID) & "'"
		End If

		If sql <> "" Then
			rs.Open sql, m_objConn, 0, 3
			If Not rs.EOF Then
				GetRecordSet(rs)
				GetMessage = True
			End If
			rs.Close
		End If
	End Function
	
	Public Function hyperlink(txt)
	  dim workingText
	  dim matches
	  dim nHits, gRE
	  dim thisMatchBefore, thisMatchAfter
	  Dim objMatch
	  Dim RE_TCP_IPT_ADDR_NAME
	  Dim RE_TCP_IPT_ADDR_IP
	  Dim RE_TCP_IPT_ADDR
	  Dim RE_EMAIL_ADDR
	  Dim RE_UNIX_PATH
	  Dim RE_WWW_URL
	  Dim RE_HREF
	  Dim RE_HTTP_URL
	  
	  RE_TCP_IPT_ADDR_NAME = "[\w\-]+([\.?=]+[\w\-]+)*"
	  RE_TCP_IPT_ADDR_IP = "[0-9]{0,3}\.[0-9]{0,3}\.[0-9]{0,3}\.[0-9]{0,3}"
	  RE_TCP_IPT_ADDR = "(" & RE_TCP_IPT_ADDR_IP & "|" & RE_TCP_IPT_ADDR_NAME & ")"
	  RE_EMAIL_ADDR = "([_a-zA-Z0-9\-\.]+@(" & RE_TCP_IPT_ADDR & "))"
	  RE_UNIX_PATH = "(/[_a-zA-Z0-9\-\.,@?^=%&:/~\+#;]+)+"
	  RE_WWW_URL = "((([^:][^/][^/])|^)([Ww][Ww][Ww])(\.(" & RE_TCP_IPT_ADDR & ")(:[0-9]+)?(" & RE_UNIX_PATH & ")*))"
	  RE_HREF = "([Ff][Tt][Pp]|[Ff][Tt][Pp][Ss]|[Hh][Tt][Tt][Pp]|[Hh][Tt][Tt][Pp][Ss])"
	  RE_HTTP_URL = "([^""]|^)(" & RE_HREF & "://([Ww][Ww][Ww]\.)?(" & RE_TCP_IPT_ADDR & ")(:[0-9]+)?(" & RE_UNIX_PATH & ")*)([^""]|$)"

	  workingText = txt
	  set gRE=server.createobject("VBScript.Regexp")
	  gRE.Global=True
	  'gRE.Pattern = "([^A-Za-z0-9‰ˆ¸ƒ÷‹'])((http://|https://|ftp://|mailto:|news:)[^\s\&lt;\&gt;\(\)\[\]]+)"
	  'gRE.Pattern="(?!&lt;.*?)(http://|www\.)([^ \r]+)(?!.*?&gt;)"
      'workingText=gRE.replace(workingText,"<a href='$1$2'>$1$2</a>")
      gRE.Pattern = RE_WWW_URL
      workingText= gRE.Replace(workingText, "$2<a href=""http://$4$5""Target=""_BLANK"">$4$5</a>")

      gRE.Pattern = RE_HTTP_URL
      workingText= gRE.Replace(workingText, "$1<a href=""$2""Target=""_BLANK"">$2</a>$11")
	
	  hyperlink = workingText
	  Set gRE = nothing
	End Function

End Class
%>