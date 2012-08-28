<%
Class CSearchManager
	Public m_intID
	Public m_strURL
	Public m_intTypeID
	Public m_dateLastIndexed

	Public m_objConn
	Public m_objSearchParser

	Public Sub Class_Initialize
		m_intID					= 0
		m_dateLastIndexed		= ""
		m_strURL					= ""

		Set m_objConn	= Server.CreateObject("ADODB.Connection")
		m_objConn.Open strDB

		Set m_objSearchParser = New CSearchParser
	End Sub

	Public Sub Class_Terminate
		m_objConn.Close
		Set m_objConn			= Nothing
		Set m_objSearchParser	= Nothing
	End Sub

	Public Sub GetData()
		Dim sql
		If m_intID > 0 Then
			sql = "SELECT TOP 1 * FROM URLs WHERE ID = " & m_intID
		End If

		If sql <> "" Then
			rs.Open sql, m_objConn, 3, 3
			If Not rs.EOF Then
				GetRecordSet(rs)
			Else
				m_intID 		= 0
			End If
			rs.Close
		End If
	End Sub

	Public Sub GetRecordSet(rs)
		m_intID					= Trim(rs("ID"))
		m_dateLastIndexed		= rs("LastIndexed")
		m_strURL				= Trim(rs("URL"))
	End Sub

	Private Sub GetFormData()
		m_intID		= Trim(Request("ID"))
		m_strURL	= Trim(Request("URL"))
	End Sub

	Private Sub PrepareDataForWriting()
		m_strURL			= Replace(m_strURL, "'", "''")
		m_strURL			= Replace(m_strURL, """", "&quot;")
		m_strURL			= Replace(m_strURL, Chr(10), " ")
		m_strURL			= Replace(m_strURL, Chr(13), " ")
	End Sub

	Public Sub Update()
		Call GetFormData()
		Call PrepareDataForWriting()
		Call m_objSearchParser.Update(m_intID, SEARCH_TYPE_GENERAL, m_strURL, "")
	End Sub

	Public Sub Insert()
		Call GetFormData()
		Call PrepareDataForWriting()

		m_intID = m_objSearchParser.Insert(SEARCH_TYPE_GENERAL, m_strURL, "")
	End Sub

	Public Sub Delete()
		Dim sql
		Dim i
		Call GetFormData()

		Call m_objSearchParser.ClearURL(m_strURL, SEARCH_TYPE_GENERAL)
	End Sub

	Public Sub OutputIDList()
		Dim sql

		Response.Write("<select name=""IDList"">")
		sql = "SELECT ID, URL FROM URLs WHERE TypeID=" & SEARCH_TYPE_GENERAL & " And Deleted=0 ORDER BY ID"
		rs.Open sql, m_objConn, 0, 3
		Do While Not rs.EOF
			Call AddSelectOption(rs("ID"), rs("ID") & " - " & rs("URL"), m_intID)
			rs.MoveNext
		Loop
		rs.Close
		Response.Write("</select>")
	End Sub

	Sub AddSelectOption(strValue, strText, strSelectedValue)
		' Adds an entry into an HTML <SELECTion> box
		Response.Write("<option ")
		If Trim(CStr(strValue)) = Trim(CStr(strSelectedValue)) Then
			Response.Write("selected ")
		End If
		Response.Write("value=""" & strValue & """>" & strText)
	End Sub

	Private Function NiceDate(dateDate)
		If IsDate(dateDate) Then
			NiceDate = Day(dateDate) & "/" & UCase(MonthName(Month(dateDate), True)) & "/" & Right(Year(dateDate), 4)
			NiceDate = NiceDate & " " & Right("0" & Hour(dateDate), 2)
			NiceDate = NiceDate & ":" & Right("0" & Minute(dateDate), 2)
			'Ignore seconds
			'NiceDate = NiceDate & ":" & Right("0" & Second(dateDate), 2)
		Else
			NiceDate = dateDate
		End If
	End Function
End Class
%>