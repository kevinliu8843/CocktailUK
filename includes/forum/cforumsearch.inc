<%
Const FORUM_TITLES_AND_MESSAGES	= "A"
Const FORUM_TITLES_ONLY			= "B"
Const FORUM_MESSAGES_ONLY		= "C"

Class CForumSearch
	Public m_objConn
	Private m_objForum

	Public m_intForumID
	Public m_strQuery		' The textual query that the user has entered
	Public m_strOption		' Title search, message search or both

	Public m_strThreadID
	Public m_intNumMessages
	Public m_strUsername
	Public m_dateThreadUpdated
	Public m_strSubject

	Public m_intOrdering

	Public m_intPageSize
	Public m_intCurrentPage

	Public Sub Class_Initialize
		Set m_objForum	= New CForum
		Set m_objConn	= Server.CreateObject("ADODB.Connection")
		m_objConn.Open(strDB)

		m_intForumID	= 0
		m_strQuery		= ""
		m_strOption		= ""

		m_intOrdering				= DATE_ORDER Or DESCENDING		' This will be stored in the cookie

		' Paging variables
		m_intPageSize				= 25
		m_intCurrentPage			= 1
	End Sub

	Public Sub Class_Terminate
		Set m_objForum	= Nothing
		Set m_objConn	= Nothing
	End Sub

	Public Function GetRecordSet(rs)
		GetRecordSet = True
		m_strThreadID		= rs("ThreadID")
		m_intNumMessages	= rs("NumMessages")
		m_dateThreadUpdated	= rs("ThreadUpdated")	' The last time that anyone posted a message to this thread
		m_strUsername		= rs("Username")
		m_strSubject		= rs("Subject")
	End Function

	Private Function GetSearchWordList(strPhrase)
		' returns only the actual words from the string entered (omits logical operators)
		GetSearchWordList = UCase(strPhrase)
		GetSearchWordList = Replace(GetSearchWordList, " AND ", ",")
		GetSearchWordList = Replace(GetSearchWordList, " OR ", ",")
		GetSearchWordList = Replace(GetSearchWordList, " NOT ", ",")
	End Function

	Private Function GetBooleanQuery(strPhrase)
		Dim i, j, intOperator, strKeyWord

		Dim aryOperators(2)
		Dim aryOperatorSQL(2)

		aryOperators(0)	= " AND "
		aryOperatorSQL(0) = "%' AND //FIELD// LIKE '%"
		aryOperators(1) = " OR "
		aryOperatorSQL(1) = "%' OR //FIELD// LIKE '%"
		aryOperators(2) = " NOT "
		aryOperatorSQL(2) = "%' AND //FIELD// NOT LIKE '%"

		GetBooleanQuery = ""
		GetBooleanQuery = GetBooleanQuery & "//FIELD// LIKE '%"
		j = 1
		For i = 1 To Len(strPhrase) + 1
			If i > j Then
				strKeyWord = Mid(strPhrase, j, i - j)
			End If

			For intOperator = 0 To UBound(aryOperators)
				If UCase(Mid(strPhrase, i, Len(aryOperators(intOperator)))) = aryOperators(intOperator) Then

					GetBooleanQuery = GetBooleanQuery & Replace(strKeyWord, "'", "''") & aryOperatorSQL(intOperator)
					j = i + Len(aryOperators(intOperator))
					Exit For
				End If
			Next
		Next
		GetBooleanQuery = GetBooleanQuery & Replace(strKeyWord, "'", "''") & "%'"
	End Function

	Public Sub FindThreads()
		Dim bRecordsToDisplay, intTotalPages, strPageControls, i
		Dim strSearchWordList

		If m_strQuery <> "" Then
			strSearchWordList	= GetSearchWordList(m_strQuery)
			strQueryinSQL		= GetBooleanQuery(m_strQuery)

			sql = "SELECT ThreadID, Subject, NumMessages, Username, ThreadUpdated FROM ForumMessages WHERE ThreadID IN("
			
			sql = sql & "SELECT DISTINCT ThreadID FROM ForumMessages WHERE "
			If m_strOption = FORUM_TITLES_AND_MESSAGES Or m_strOption = FORUM_TITLES_ONLY Then
				sql = sql & "(" & Replace(strQueryinSQL, "//FIELD//", "Subject") & ") "
			End If

			If m_strOption = FORUM_TITLES_AND_MESSAGES Then sql = sql & " OR "

			If m_strOption = FORUM_TITLES_AND_MESSAGES Or m_strOption = FORUM_MESSAGES_ONLY Then
				sql = sql & "(" & Replace(strQueryinSQL, "//FIELD//", "Message") & ") "
			End If
			sql = sql & " And ForumID = " & m_intForumID
			sql = sql & " And Status = 1) AND ParentID IS NULL And Status = 1 And ForumID = " & m_intForumID & " "

			If m_intOrdering And SUBJECT_ORDER Then
				sql = sql & "ORDER BY Subject"
			ElseIf m_intOrdering And POSTER_ORDER Then
				sql = sql & "ORDER BY Username"
			ElseIf m_intOrdering And POSTS_ORDER Then
				sql = sql & "ORDER BY NumMessages"
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
		Else
			bRecordsToDisplay = False
		End If

		If bRecordsToDisplay = True Then
			rs.AbsolutePage = m_intCurrentPage

			intTotalPages	= rs.PageCount
			If intTotalPages > 1 Then
				strPageControls = GetPageControls(intTotalPages)
				Response.Write("<center>" & strPageControls & "<br></center>")
			End If
		End If

		If bRecordsToDisplay = True Then
			Response.Write("<table width=""100%"" cellspacing=""1"" cellpadding=""2"">")
			Call DisplayHeader()
		Else
			Response.Write("<hr size=""1"" noshade style=""color: Silver;"">")
			Response.Write("<font color=""red"">&#149;</font>&nbsp;No Subjects Found")
		End If

		If m_strQuery <> "" Then
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
				<%= m_objForum.MessageHREF(m_intForumID, m_strThreadID, "", 1, strSearchWordList, False) & m_strSubject%></a>&nbsp;
				</td>
				<td valign="top"><%=m_intNumMessages%>&nbsp;</td>
				<td valign="top"><%=m_strUsername%>&nbsp;</td>
				<td align="left" valign="top"><%= Replace(m_objForum.FriendlyDate(m_dateThreadUpdated)," ", "&nbsp;")%>&nbsp;</td>
				<td align="left" valign="top"><%= FormatDateTime(m_dateThreadUpdated, 4)%></td>
				</tr>
				<%
				rs.MoveNext
			Next
			rs.Close
		End If

		If bRecordsToDisplay = True Then
			%>
			<tr><td colspan="5" valign="top"><hr size="1" noshade style="color: Silver;"></td></tr>
			</table>
			<%
		End If
	End Sub

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
			Response.Write( m_objForum.SearchHREF(m_intForumID, "AS", m_intCurrentPage, m_strQuery, m_strOption) & "Subject</a>")
		End If

		If (m_intOrdering AND SUBJECT_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.SearchHREF(m_intForumID, "AS", m_intCurrentPage, m_strQuery, m_strOption) & strUpImg & "</a>")
			Else
				Response.Write(m_objForum.SearchHREF(m_intForumID, "DS", m_intCurrentPage, m_strQuery, m_strOption) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>
		<td bgcolor="#ebe2f7">
		<%
		If (m_intOrdering AND POSTS_ORDER) Then 
			Response.Write("<b>Posts</b>&nbsp;") 
		Else
			Response.Write(m_objForum.SearchHREF(m_intForumID, "AN", m_intCurrentPage, m_strQuery, m_strOption) & "Posts</a>&nbsp;")
		End If

		If (m_intOrdering AND POSTS_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.SearchHREF(m_intForumID, "AN", m_intCurrentPage, m_strQuery, m_strOption) & strUpImg & "</a>&nbsp;")
			Else
				Response.Write(m_objForum.SearchHREF(m_intForumID, "DN", m_intCurrentPage, m_strQuery, m_strOption) & strDownImg & "</a>&nbsp;")
			End If
		End If
		%>
		</td>
		
		<td bgcolor="#ebe2f7">
		<% 
		If (m_intOrdering AND POSTER_ORDER) Then 
			Response.Write("<b>Poster</b>") 
		Else
			Response.Write(m_objForum.SearchHREF(m_intForumID, "AP", m_intCurrentPage, m_strQuery, m_strOption) & "Poster</a>")
		End If

		If (m_intOrdering AND POSTER_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.SearchHREF(m_intForumID, "AP", m_intCurrentPage, m_strQuery, m_strOption) & strUpImg & "</a>")
			Else
				Response.Write(m_objForum.SearchHREF(m_intForumID, "DP", m_intCurrentPage, m_strQuery, m_strOption) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>		

		<td bgcolor="#ebe2f7" colspan="2">
		<% 
		If (m_intOrdering AND DATE_ORDER) Then 
			Response.Write("<b>Last&nbsp;Reply</b>") 
		Else
			Response.Write(m_objForum.SearchHREF(m_intForumID, "AD", m_intCurrentPage, m_strQuery, m_strOption) & "Last&nbsp;Reply</a>")
		End If

		If (m_intOrdering AND DATE_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(m_objForum.SearchHREF(m_intForumID, "AD", m_intCurrentPage, m_strQuery, m_strOption) & strUpImg & "</a>")
			Else
				Response.Write(m_objForum.SearchHREF(m_intForumID, "DD", m_intCurrentPage, m_strQuery, m_strOption) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>
		</tr>
<%
	End Sub

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
			Case Else	m_intOrdering = (m_intOrdering Or DATE_ORDER)
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
		ElseIf (m_intOrdering And POSTS_ORDER) Then
			GetOrder = GetOrder & "N"
		Else
			GetOrder = GetOrder & "D"
		End If
	End Function

	Public Sub SaveSettings()
		Response.Cookies(FORUM_COOKIE_FILENAME)(FORUM_THREAD_ORDERING)	= m_intOrdering
	End Sub

	Public Sub LoadSettings()
		If Request.Cookies(FORUM_COOKIE_FILENAME).HasKeys Then
			m_intOrdering		= Request.Cookies(FORUM_COOKIE_FILENAME)(FORUM_THREAD_ORDERING)
		End If
	End Sub

	Private Function GetPageControls(intTotalPages)
		Dim i

		If Int(m_intCurrentPage) > 1 Then
			GetPageControls = GetPageControls & m_objForum.SearchHREF(m_intForumID, GetOrder(), Int(m_intCurrentPage)-1, m_strQuery, m_strOption) & "&lt;&lt; Previous Page</a>&nbsp;"
		Else
			GetPageControls = GetPageControls & "&lt;&lt; Previous Page&nbsp;"
		End If

		For i=1 to intTotalPages
			If Int(i) = Int(m_intCurrentPage) Then
				GetPageControls = GetPageControls & "<b>" & i & "</b>&nbsp;"
			Else
				GetPageControls = GetPageControls & m_objForum.SearchHREF(m_intForumID, GetOrder(), i, m_strQuery, m_strOption) & i & "</a>&nbsp;"
			End If
		Next

		If Int(m_intCurrentPage) < Int(intTotalPages) Then
			GetPageControls = GetPageControls & m_objForum.SearchHREF(m_intForumID, GetOrder(), Int(m_intCurrentPage)+1, m_strQuery, m_strOption) & "Next Page &gt;&gt;</a>"
		Else
			GetPageControls = GetPageControls & "Next Page &gt;&gt;"
		End If
	End Function


End Class
%>