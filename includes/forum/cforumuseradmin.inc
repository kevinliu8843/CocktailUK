<%
Class CForumUserAdmin
	Public m_objConn
	Private m_objForum
	Public m_intOrdering

	' Paging variables
	Public m_intPageSize			' The number of messages to display per page
	Public m_intCurrentPage			'

	Public Sub Class_Initialize
		Set m_objForum	= New CForum
		Set m_objConn	= Server.CreateObject("ADODB.Connection")
		m_objConn.Open(strDB)

		m_intOrdering				= USERNAME_ORDER Or ASCENDING

		' Paging variables
		m_intPageSize				= 20
		m_intCurrentPage			= 1
	End Sub
	
	Public Sub Class_Terminate
		Set m_objForum	= Nothing
		Set m_objConn	= Nothing
	End Sub

	Function IsModerator
		IsModerator = (intForumStatus=1)
	End Function

	Private Function GetPageControls(intTotalPages)
		Dim i

		If Int(m_intCurrentPage) > 1 Then
			GetPageControls = GetPageControls & UserAdminLink(GetOrder(), Int(m_intCurrentPage)-1) & "&lt;&lt; Previous Page</a>&nbsp;"
		Else
			GetPageControls = GetPageControls & "&lt;&lt; Previous Page&nbsp;"
		End If

		For i=1 to intTotalPages
			If Int(i) = Int(m_intCurrentPage) Then
				GetPageControls = GetPageControls & "<b>" & i & "</b>&nbsp;"
			Else
				GetPageControls = GetPageControls & UserAdminLink(GetOrder(), i) & i & "</a>&nbsp;"
			End If

			If (i Mod 25)=0 Then
				GetPageControls = GetPageControls & "<br>"
			End If
		Next

		If Int(m_intCurrentPage) < Int(intTotalPages) Then
			GetPageControls = GetPageControls & UserAdminLink(GetOrder(), Int(m_intCurrentPage)+1) & "Next Page &gt;&gt;</a>"
		Else
			GetPageControls = GetPageControls & "Next Page &gt;&gt;"
		End If
	End Function

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Private Function UserAdminLink(strOrder, intPageNumber)
		UserAdminLink = "<a href=""/forums/edit_users.asp?order=" & strOrder
		If intPageNumber <> "" Then
			UserAdminLink = UserAdminLink & "&pg=" & intPageNumber
		End If

		UserAdminLink = UserAdminLink & """>"
	End Function

	Private Sub DisplayHeader()
		Dim strUpImg, strDownImg
		strUpImg	= "<img src=""/images/forums/up.gif"" width=""11"" height=""11"" border=""0"">"
		strDownImg	= "<img src=""/images/forums/down.gif"" width=""11"" height=""11"" border=""0"">"
%>
		<tr>
			<td colspan="6"><hr size="1" noshade style="color: Silver;"></td>
		</tr>
		<tr>
		<td>
			&nbsp;
		</td>

		<td width="25%">
		<% 
		If (m_intOrdering AND USERNAME_ORDER) Then 
			Response.Write("<b>Username</b>") 
		Else
			Response.Write(UserAdminLink("AU", m_intCurrentPage) & "Username</a>")
		End If

		If (m_intOrdering AND USERNAME_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(UserAdminLink("AU", m_intCurrentPage) & strUpImg & "</a>")
			Else
				Response.Write(UserAdminLink("DU", m_intCurrentPage) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>		

		<td width="25%">
		<% 
		If (m_intOrdering AND FIRSTNAME_ORDER) Then 
			Response.Write("<b>Firstname</b>") 
		Else
			Response.Write(UserAdminLink("AF", m_intCurrentPage) & "Firstname</a>")
		End If

		If (m_intOrdering AND FIRSTNAME_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(UserAdminLink("AF", m_intCurrentPage) & strUpImg & "</a>")
			Else
				Response.Write(UserAdminLink("DF", m_intCurrentPage) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>		

		<td width="25%">
		<% 
		If (m_intOrdering AND LASTNAME_ORDER) Then 
			Response.Write("<b>Lastname</b>") 
		Else
			Response.Write(UserAdminLink("AL", m_intCurrentPage) & "Lastname</a>")
		End If

		If (m_intOrdering AND LASTNAME_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(UserAdminLink("AL", m_intCurrentPage) & strUpImg & "</a>")
			Else
				Response.Write(UserAdminLink("DL", m_intCurrentPage) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>		

		<td width="25%">
		<% 
		If (m_intOrdering AND EMAIL_ORDER) Then 
			Response.Write("<b>Email</b>") 
		Else
			Response.Write(UserAdminLink("AE", m_intCurrentPage) & "Email</a>")
		End If

		If (m_intOrdering AND EMAIL_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(UserAdminLink("AE", m_intCurrentPage) & strUpImg & "</a>")
			Else
				Response.Write(UserAdminLink("DE", m_intCurrentPage) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>		

		<td>
		<% 
		If (m_intOrdering AND STATUS_ORDER) Then 
			Response.Write("<b>Status</b>") 
		Else
			Response.Write(UserAdminLink("AS", m_intCurrentPage) & "Status</a>")
		End If

		If (m_intOrdering AND STATUS_ORDER) Then 
			If (m_intOrdering AND DESCENDING) Then
				Response.Write(UserAdminLink("AS", m_intCurrentPage) & strUpImg & "</a>")
			Else
				Response.Write(UserAdminLink("DS", m_intCurrentPage) & strDownImg & "</a>")
			End If
		End If
		%>
		</td>		

		</tr>
		<tr>
			<td colspan="6"><hr size="1" noshade style="color: Silver;"></td>
		</tr>
<%
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

		If (m_intOrdering And USERNAME_ORDER) Then
			GetOrder = GetOrder & "U"
		ElseIf (m_intOrdering And STATUS_ORDER) Then
			GetOrder = GetOrder & "S"
		ElseIf (m_intOrdering And FIRSTNAME_ORDER) Then
			GetOrder = GetOrder & "F"
		ElseIf (m_intOrdering And LASTNAME_ORDER) Then
			GetOrder = GetOrder & "L"
		Else
			GetOrder = GetOrder & "E"
		End If
	End Function

'**********************************************************************************************************
'**********************************************************************************************************
'**********************************************************************************************************

	Public Sub SetOrder(strOrder)
		' strOrder will be a 2 byte string 
		' (1st character = direction)
		' (2nd character = sort type)
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
			Case "U"	m_intOrdering = (m_intOrdering Or USERNAME_ORDER)
			Case "S"	m_intOrdering = (m_intOrdering Or STATUS_ORDER)
			Case "F"	m_intOrdering = (m_intOrdering Or FIRSTNAME_ORDER)
			Case "L"	m_intOrdering = (m_intOrdering Or LASTNAME_ORDER)
			Case Else	m_intOrdering = (m_intOrdering Or EMAIL_ORDER)
		End Select
	End Sub
End Class
%>