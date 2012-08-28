<%
Class CForumThread
	Public m_objConn, m_objConnEdit				' The connection string
	Private m_objForum

	' Member variables that are contained within the ForumMessages table ----
	Public m_strID					' ID of this Message
	Public m_strParentID			' ID of the direct parent.  Can be NULL
	Public m_strThreadID			' ID of the first ancestor.  i.e Parent of Parent...
	Public m_intUID					' ID of the user who submitted this message
	Public m_intForumID				' 
	Public m_intNumChildren			' Number of direct children this message has
	Public m_intNumMessages			' Number of total messages this message has beneath it (children, grandchildren, etc.)
	Public m_intStatus				' 0 = Delete, 1 = Normal
	Public m_datePosted				'
	Public m_dateThreadUpdated		'
	Public m_strSubject				'
	Public m_strUsername			' Copy of Username from Members database (for quickness)
	Public m_strMessage				'
	Public m_intViewed
	Public m_bShowControls 			' Show page controls logical

	Public m_intStoryID				' Is this thread related to a story item?
	'----------------------------------------------------------------

	Public m_intOrdering

	Private m_bHighlight				' Internal 'global' variable, used to toggle displaying of background colour or not

	' Paging variables
	Public m_intPageSize				' The number of threads to display per page
	Public m_intCurrentPage			'

	Public m_bIsEditing				' True if we are in editing mode

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Public Sub Class_Initialize
		Set m_objForum	= New CForum
		set m_objConn	= Server.CreateObject("ADODB.Connection")
		m_objConn.Open(strDB)
		set m_objConnEdit	= Server.CreateObject("ADODB.Connection")
		m_objConnEdit.Open(strDBMod)

		' Clear off database fields ------------------
		m_strID				= ""
		m_strParentID		= ""
		m_strThreadID		= ""
		m_intUID			= 0
		m_intForumID		= 0
		m_intNumChildren	= 0
		m_intNumMessages	= 0
		m_intStatus			= 0
		m_datePosted		= Now()
		m_dateThreadUpdated	= Now()
		m_strSubject		= ""
		m_strUsername		= Session("uname")
		m_strMessage		= ""
		m_intViewed			= 0
		m_bShowControls 	= True

		m_intStoryID		= 0
		'---------------------------------------------

		m_intOrdering				= DATE_ORDER Or DESCENDING		' This will be stored in the cookie

		' Paging variables
		m_intPageSize				= 25
		m_intCurrentPage			= 1

		m_bIsEditing				= False
	End Sub

	Public Sub Class_Terminate
		m_objConn.close
		m_objConnEdit.close
		Set m_objForum = Nothing
		Set m_objConn = Nothing
		Set m_objConnEdit = Nothing
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
		m_intStatus			= rs("Status")
		m_datePosted		= rs("DatePosted")
		m_dateThreadUpdated	= rs("ThreadUpdated")	' The last time that anyone posted a message to this thread
		m_strSubject		= rs("Subject")
		m_strUsername		= rs("Username")
		m_intStoryID		= rs("StoryID")
		m_intViewed			= rs("viewed")
	End Function
	

	Public Sub SaveSettings()
		Response.Cookies(FORUM_COOKIE_FILENAME)(FORUM_THREAD_ORDERING)	= m_intOrdering
	End Sub


	Public Sub LoadSettings()
		If Request.Cookies(FORUM_COOKIE_FILENAME).HasKeys Then
			m_intOrdering		= Request.Cookies(FORUM_COOKIE_FILENAME)(FORUM_THREAD_ORDERING)
		End If
	End Sub
	
	' Displays the threads for the current forum
	Public Function Display(rs)
		' Returns True if this actually displayed any threads

		Dim strPageControls, intTotalPages, i
		Dim sql, bRecordsToDisplay

		sql = "SELECT ID, ParentID, ThreadID, ForumID, NumChildren, NumMessages, StoryID, Status, DatePosted, ThreadUpdated, Subject, Username, viewed "
		sql = sql & "FROM ForumMessages WHERE ForumID = " & m_intForumID & " AND ParentID IS NULL AND Status = 1 "


		If m_intOrdering And SUBJECT_ORDER Then
			sql = sql & "ORDER BY Subject"
		ElseIf m_intOrdering And POSTER_ORDER Then
			sql = sql & "ORDER BY Username"
		ElseIf m_intOrdering And POSTS_ORDER Then
			sql = sql & "ORDER BY NumMessages"
		ElseIf m_intOrdering And VIEWED_ORDER Then
			sql = sql & "ORDER BY viewed"
		Else
			sql = sql & "ORDER BY ThreadUpdated"
		End If
		If m_intOrdering And DESCENDING Then
			sql = sql & " DESC"
		End If

		' Handle pages ------------------------------------------------------------------------
		rs.PageSize			= m_intPageSize
		rs.CursorLocation	= 2
		rs.CacheSize		= m_intPageSize * 10

		rs.Open sql, m_objConn, 3, 3
		bRecordsToDisplay = (NOT rs.EOF)
		Display = bRecordsToDisplay
	
		If bRecordsToDisplay = True Then
			rs.AbsolutePage = m_intCurrentPage

			intTotalPages	= rs.PageCount
			If intTotalPages > 1 AND m_bShowControls Then
				strPageControls = GetPageControls(intTotalPages)
%>
		<table cellpadding="0" cellspacing="0" border="0" width="100%" bgcolor="#ebe2f7">
		<tr>
			<td><font color="red">&#149;</font>&nbsp;<a href="/forums/search.asp?fid=<%=m_intForumID%>">Search</a></td>
			<td width="99%" align="center"><%= strPageControls%></td>
			<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
		</table>
<%
			ElseIf m_bShowControls Then
%>
		<table cellpadding="0" cellspacing="0" border="0" width="100%">
		<tr>
			<td><font color="red">&#149;</font>&nbsp;<a href="/forums/search.asp?fid=<%=m_intForumID%>">Search</a></td>
			<td width="99%" align="center"><%=strPageControls%></td>
		</tr>
		</table>
<%
			End If
		End If

		If m_bShowControls  Then Response.Write("<br>")

		If m_bIsEditing = True Then
		%>		
		<form action="<%=m_objForum.ThreadLink(m_intForumID, GetOrder(), m_intCurrentPage, m_bIsEditing)%>" method="post">
		<input type="submit" name="action" value="Delete">
		<%
		End If

		If bRecordsToDisplay = True Then
			%>
			<table width="100%" cellspacing="1" cellpadding="2">
			<%
			Call DisplayHeader()
		Else
		%>
		<a href="/forums/new_subject.asp?fid=<%=m_intForumID%>">New Subject</a><br>
		<%
		End If

		For i=1 to m_intPageSize
			If RS.EOF Then
				Exit For
			End If

			GetRecordSet(rs)
			%>
			<%If (i Mod 2) = 1 Then %>
			<tr>
			<%Else%>
			<tr bgcolor="#F0F0F0">
			<%End If%>
			<td valign="top">
			<% If m_bIsEditing = True Then %>
			<input type="checkbox" name="ID" value="<%=m_objForum.CompactID(m_strID)%>">
			<% End If %>
			<%= m_objForum.MessageHREF(m_intForumID, m_strID, "", 1, "", m_bIsEditing) & m_strSubject%></a>&nbsp;
			</td>
			<td valign="top"><%=m_intNumMessages%>&nbsp;</td>
			<td valign="top"><%= m_intViewed%>&nbsp;</td>
			<td valign="top"><%= m_strUsername%>&nbsp;</td>
			<td align="left" valign="top"><%= Replace(m_objForum.FriendlyDate(m_dateThreadUpdated)," ", "&nbsp;")%>&nbsp;</td>
			<td align="left" valign="top"><%= FormatDateTime(m_dateThreadUpdated, 4)%></td>
			</tr>
			<%
			rs.MoveNext
		Next
		rs.Close

		If bRecordsToDisplay = True Then
			%>
			</table>
			<%
		End If

		If m_bIsEditing = True Then
			%></form><%
		End If
	End Function

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Private Sub DisplayHeader()
		Dim strUpImg, strDownImg
		strUpImg	= "<img src=""/images/forums/u.gif"" width=""14"" height=""14"" border=""0"">"
		strDownImg	= "<img src=""/images/forums/d.gif"" width=""14"" height=""14"" border=""0"">"
		%>
		<tr>
		<td bgcolor="#ebe2f7">
		<% 
		If (m_intOrdering AND SUBJECT_ORDER) Then 
			Response.Write("<b>Subject</b>") 
		Else
			Response.Write( m_objForum.ThreadHREF(m_intForumID, "AS", m_intCurrentPage, m_bIsEditing) & "Subject</a>")
		End If

		If (m_intOrdering AND SUBJECT_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "AS", m_intCurrentPage, m_bIsEditing) & strUpImg & "</a>")
			Else
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "DS", m_intCurrentPage, m_bIsEditing) & strDownImg & "</a>")
			End If
		End If
		%>
		(<a href="/forums/new_subject.asp?fid=<%=m_intForumID%>">New</a>)
		</td>

		<td bgcolor="#ebe2f7">
		<%
		If (m_intOrdering AND POSTS_ORDER) Then 
			Response.Write("<b>Posts</b>&nbsp;") 
		Else
			Response.Write(m_objForum.ThreadHREF(m_intForumID, "AN", m_intCurrentPage, m_bIsEditing) & "Posts</a>&nbsp;")
		End If

		If (m_intOrdering AND POSTS_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "AN", m_intCurrentPage, m_bIsEditing) & strUpImg & "</a>&nbsp;")
			Else
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "DN", m_intCurrentPage, m_bIsEditing) & strDownImg & "</a>&nbsp;")
			End If
		End If
		%></td><td bgcolor="#ebe2f7"><%
		If (m_intOrdering AND VIEWED_ORDER) Then 
			Response.Write("<b>Views</b>&nbsp;") 
		Else
			Response.Write(m_objForum.ThreadHREF(m_intForumID, "AV", m_intCurrentPage, m_bIsEditing) & "Views</a>&nbsp;")
		End If

		If (m_intOrdering AND VIEWED_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "AV", m_intCurrentPage, m_bIsEditing) & strUpImg & "</a>&nbsp;")
			Else
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "DV", m_intCurrentPage, m_bIsEditing) & strDownImg & "</a>&nbsp;")
			End If
		End If
		%></td><td bgcolor="#ebe2f7"><% 
		If (m_intOrdering AND POSTER_ORDER) Then 
			Response.Write("<b>Poster</b>") 
		Else
			Response.Write(m_objForum.ThreadHREF(m_intForumID, "AP", m_intCurrentPage, m_bIsEditing) & "Poster</a>")
		End If

		If (m_intOrdering AND POSTER_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "AP", m_intCurrentPage, m_bIsEditing) & strUpImg & "</a>")
			Else
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "DP", m_intCurrentPage, m_bIsEditing) & strDownImg & "</a>")
			End If
		End If
		%></td><td bgcolor="#ebe2f7" colspan="2"><% 
		If (m_intOrdering AND DATE_ORDER) Then 
			'Response.Write("<b>Date</b>") 
			Response.Write("<b>Last&nbsp;Reply</b>") 
		Else
			'Response.Write(m_objForum.ThreadHREF(m_intForumID, "AD", m_intCurrentPage, m_bIsEditing) & "Date</a>")
			Response.Write(m_objForum.ThreadHREF(m_intForumID, "AD", m_intCurrentPage, m_bIsEditing) & "Last&nbsp;Reply</a>")
		End If

		If (m_intOrdering AND DATE_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "AD", m_intCurrentPage, m_bIsEditing) & strUpImg & "</a>")
			Else
				Response.Write(m_objForum.ThreadHREF(m_intForumID, "DD", m_intCurrentPage, m_bIsEditing) & strDownImg & "</a>")
			End If
		End If
		%></td></tr><%
	End Sub

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Public Sub SetOrder(strOrder)
		' strOrder will be a 2 byte string 
		' (1st character = direction)
		' (2nd character = sort type, i.e.)
		If Len(strOrder) <> 2 Then
			Exit Sub
		End If

		Select Case Mid(strOrder, 1, 1)
			Case "A"
				m_intOrdering = ASCENDING
			Case Else
				m_intOrdering = DESCENDING
		End Select

		Select Case Mid(strOrder, 2, 1)
			Case "S"	m_intOrdering = (m_intOrdering Or SUBJECT_ORDER)
			Case "P"	m_intOrdering = (m_intOrdering Or POSTER_ORDER)
			Case "N"	m_intOrdering = (m_intOrdering Or POSTS_ORDER)
			Case "V"	m_intOrdering = (m_intOrdering Or VIEWED_ORDER)
			Case Else	m_intOrdering = (m_intOrdering Or DATE_ORDER)
		End Select
	End Sub

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

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
		ElseIf (m_intOrdering And POSTS_ORDER) Then
			GetOrder = GetOrder & "N"
		Else
			GetOrder = GetOrder & "D"
		End If
	End Function

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Public Function AddThread(rs)
		Dim sql, strThreadID

		m_strUsername = Session("uname")

		sql = "INSERT INTO ForumMessages (ParentID, ThreadID, ForumID, NumChildren, NumMessages, StoryID, Status, ReIndex, DatePosted, ThreadUpdated, Subject, Username, Message)"
'		sql = sql & "VALUES (newID(),"
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

		m_strMessage	= m_objForum.StripHTMLSpecial(m_strMessage)
		m_strMessage	= m_objForum.FilterALLSwearWords( m_strMessage, False )
		m_strMessage	= Replace(m_strMessage, "'", "''")
		m_strMessage	= Replace(m_strMessage, VBCrLf, "<br>")


		m_datePosted = Day(m_datePosted) & "/" & MonthName(Month(m_datePosted)) & "/" & Year(m_datePosted) & " " & Time()


		If m_strSubject <> "" And m_strMessage <> "" Then
			sql = sql & m_intForumID		& ","
			sql = sql & m_intNumChildren	& ","
			sql = sql & m_intNumMessages	& ","
			sql = sql & m_intStoryID		& ","
			sql = sql & "1, " ' Status
			sql = sql & "1, " ' ReIndex
			sql = sql & "'" & m_datePosted	& "',"
			sql = sql & "'" & m_datePosted	& "',"
			sql = sql & "'" & m_strSubject	& "',"
			sql = sql & "'" & m_strUsername & "',"
			sql = sql & "'" & m_strMessage  & "'"
			sql = sql & ")"

			m_objConnEdit.Execute(sql)

			'------------------------------------------------------------------------------------------------
			'This is a top-level message, set the ThreadID to the the same as the ID of the new record
			'Find the ID of the message just INSERTed

			sql = "SELECT TOP 1 ID, ThreadID FROM ForumMessages WHERE "
			sql = sql & "DatePosted = '"	& m_datePosted & "' AND "
			sql = sql & "Subject = '"		& m_strSubject & "' AND "
			sql = sql & "Username = '"		& m_strUsername & "'"

			rs.Open sql, m_objConn, 3, 3
			If NOT rs.EOF Then
				m_objConnEdit.eXECUTE("UPDATE ForumMessages SET ThreadID=ID WHERE ID='"&rs("ID")&"'")
				m_strThreadID	= m_objForum.CompactID(rs("ID"))
				rs.Update
			Else
				' This should not happen!!!
			End If
			rs.Close
			
			' Add one to the number of Threads this Forum has got
			sql = "UPDATE Forums SET NumThreads = NumThreads + 1 WHERE ID = " & m_intForumID
			m_objConnEdit.Execute(sql)

			AddThread = True
		Else
			AddThread = False
		End If

	End Function

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Private Function GetPageControls(intTotalPages)
		Dim i

		If Int(m_intCurrentPage) > 1 Then
			GetPageControls = GetPageControls & m_objForum.ThreadHREF(m_intForumID, GetOrder(), Int(m_intCurrentPage)-1, m_bIsEditing) & "&lt;&lt; Previous Page</a>&nbsp;"
		Else
			GetPageControls = GetPageControls & "&lt;&lt; Previous Page&nbsp;"
		End If

		For i=1 to intTotalPages
			If Int(i) = Int(m_intCurrentPage) Then
				GetPageControls = GetPageControls & "<b>" & i & "</b>&nbsp;"
			Else
				GetPageControls = GetPageControls & m_objForum.ThreadHREF(m_intForumID, GetOrder(), i, m_bIsEditing) & i & "</a>&nbsp;"
			End If
		Next

		If Int(m_intCurrentPage) < Int(intTotalPages) Then
			GetPageControls = GetPageControls & m_objForum.ThreadHREF(m_intForumID, GetOrder(), Int(m_intCurrentPage)+1, m_bIsEditing) & "Next Page &gt;&gt;</a>"
		Else
			GetPageControls = GetPageControls & "Next Page &gt;&gt;"
		End If
	End Function

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Public Sub DeleteThreads(rs)
		Dim strID, sql, intNumMessages

		For Each strID in Request("ID")
			' We need to know how many messages we're deleting
			sql = "SELECT TOP 1 NumMessages FROM ForumMessages WHERE ID = '" & m_objForum.ExpandID(strID) & "'"
			rs.Open sql, m_objConn, 0, 3
			If NOT rs.EOF Then
				intNumMessages = rs("NumMessages")
			Else
				intNumMessages = 0
			End If

			rs.Close

			' Now decrease the number of messages that this Forum has
			
			sql = "UPDATE Forums SET NumMessages = NumMessages - " &  intNumMessages & " WHERE ID = " & m_intForumID
			m_objConnEdit.Execute(sql)

			' We'll have one less thread too
			sql = "UPDATE Forums SET NumThreads = NumThreads - 1 WHERE ID = " & m_intForumID
			m_objConnEdit.Execute(sql)

			' Finally "kill" the thread and it's messages...
			sql = "UPDATE ForumMessages SET Status = 0, ReIndex=1 WHERE ThreadID = '" & m_objForum.ExpandID(strID) & "'"
			m_objConnEdit.Execute sql

			' Finally "kill" the parent thread
			sql = "UPDATE ForumMessages SET Status = 0, ReIndex=1 WHERE ID = '" & m_objForum.ExpandID(strID) & "'"
			m_objConnEdit.Execute sql

		Next
	End Sub
End Class
%>