<%
CONST FONT_SMALL = "<font face=""Verdana, Arial"" size=""1"">"
CONST FONT_STANDARD = "<font face=""Verdana, Arial"" size=""2"">"


' Database DataTypes (Can be combined)
Const tID		= 1
Const tNUMERIC	= 2
Const tTEXT		= 4
Const tVARCHAR	= 8
Const tDATE		= 16
Const tMUZE		= 32

' Filter Styles
Const fEQUAL	= 1		' =
Const fGT_EQ	= 2  	' >=
Const fLT_EQ	= 3  	' <=
Const fLIKE		= 4		' LIKE '%   %'

Const DEFAULT		= True
Const NOT_DEFAULT	= False

Const HOME 			= "H"
Const VIEW 			= "V"
Const SAVE 			= "S"
Const CONFIGURE 	= "C"
Const REPORT 		= "R"
Const SENDREPORT 	= "P"

'For GenerateSQL
Const SELECTEDFIELDS	= 1
Const ALLFIELDS			= 2

Class CDBViewer
	' The information stored for each of the fields we have...
	Public m_aryTitle
	Public m_aryFieldName
	Public m_aryFilter			' The filter text for this field
	Public m_aryFilterStyle		' <=, LIKE etc.
	Public m_aryInclude			' True if we are to include this field (by default) in our filter (i.e. checkbox has been clicked)
	Public m_aryDataType		' The underlying datatype for this field
	Public m_aryData			' Storage for a record's worth of data
	Public m_intNumFields		' The number of fields we have

	' --------------------------------------------------------
	Public 	m_intPageSize

	Public 	m_strOrderDirection		' Either A or D
	Public 	m_intOrder				' Will equal ID, DATEON etc.

	Public 	m_strPreferredSelectedFields	' The string of Link Characters of the fields that we want (and the order)
	Public 	m_strDefaultSelectedFields 	' The default if nothing else has been selected

	Public 	m_intCurrentPage
	Private m_bHighlight		' Used to alternate the bgcolor

	Public 	m_strDSN			' The connection we're using
	Public 	m_strTableName		' The database table we are using
	Public  m_rs
	Public  m_connConnection
	Public m_objParent

	Public m_strAction
	Public m_strApplicationName	'The name of this app eg. Perth2k, Perth2kLive, etc.
	Public Sub Class_Initialize
		' See comments above for explanation of these fields...
		Redim m_aryTitle(99)
		Redim m_aryFieldname(99)
		Redim m_aryDataType(99)
		Redim m_aryFilter(99)
		Redim m_aryFilterStyle(99)
		Redim m_aryInclude(99)
		Redim m_aryData(99)

		m_intNumFields 			= 0
		Set m_rs				= Server.CreateObject("ADODB.Recordset")
		Set m_connConnection	= Server.CreateObject("ADODB.Connection")
		Set m_objParent 		= Nothing

		m_strDefaultSelectedFields 	= "" 
		m_strAction 				= HOME
	End Sub

	Public Sub Class_Terminate
		Set m_rs				= Nothing
		Set m_connConnection	= Nothing
		Set m_objParent 		= Nothing
	End Sub

	Public Sub SetParentObject( objParent)
		Set m_objParent = objParent
	End Sub

	Public Function GetCookieName()
		Dim intStart, intEnd, strUpperCaseDSN
		' Generate cookie name based on table and dsn
		'e.g. "Perth2" + "dsn=media1_SQL2;uid=sa;pwd=arnold" 
		' becomes "PERTH2-MEDIA1_SQL2"
		GetCookieName = UCase(m_strTableName) 

		strUpperCaseDSN = UCase(m_strDSN)
		intStart = Instr(strUpperCaseDSN, "DSN=")
		If intStart > 0 Then
			intStart = intStart + 4
			intEnd = Instr(intStart, strUpperCaseDSN, ";")
			If intEnd = 0 Then intEnd = Len(strUpperCaseDSN)
			
			GetCookieName = GetCookieName & "-" & Mid(strUpperCaseDSN, intStart, intEnd - intStart)
		End If
	End Function

	Public Sub AddField(intEntry, strTitle, strFieldName, intDataType, strDefaultValue, intFilterStyle, bInclude, bIsDefault)
		m_aryTitle(intEntry)			= strTitle
		m_aryFieldName(intEntry)		= strFieldName
		m_aryDataType(intEntry)			= intDataType
		m_aryFilter(intEntry)			= strDefaultValue
		m_aryFilterStyle(intEntry)		= intFilterStyle
		m_aryInclude(intEntry)			= bInclude

		If bIsDefault Then
			m_strDefaultSelectedFields = m_strDefaultSelectedFields & GetLinkChar(intEntry)
		End If

		m_intNumFields = m_intNumFields + 1
	End Sub

	Public Sub DisplayTop()
		Dim strCookieName
		m_strAction = Request("Action")


		strCookieName = GetCookieName()
		
		If m_strAction = "" Then
			If (Request.Form("order") <> "" Or Request.QueryString("order") <> "") Then
				'We've probably clicked on a link
				m_strAction = VIEW
			Else
				m_strAction = HOME
			End If
		End If

		If m_strAction = SAVE Then
			If Request("SelectedFields") <> "" Then
				Response.Cookies(strCookieName) 			= ""
				Response.Cookies(strCookieName).Expires 	= Date + 365
				Response.Cookies(strCookieName)("Fields")	= Request("SelectedFields")
			End If
			m_strAction = HOME
		End If

		'Load cookie preferences
		m_strPreferredSelectedFields = Request.Cookies(strCookieName)("Fields")

		' Crossover check (from Old style to New Style - can kill when everyone's cookies are cool)
		If m_strPreferredSelectedFields <> "" Then
			If Not IsNumeric(Mid(m_strPreferredSelectedFields, 1, 1)) Then
				m_strPreferredSelectedFields = ""
			End If
		End If

		m_intCurrentPage		= 1
		m_intPageSize			= 20
		Call SetPage()
		Call SetOrder()
		Call GetParameters()

		Call DisplayHeader()
	End Sub

	Public Sub DisplayBottom()
		Select Case m_strAction
			Case VIEW	Call DisplayEntries()
		End Select

		Call DisplayFooter()
	End Sub

	Public Sub DisplayStandardFilter(intItemID)
		Response.Write("<tr>")
		Call DisplayStandardFilter_Title(intItemID)
		Call DisplayStandardFilter_Control(intItemID)
		Call DisplayStandardFilter_Checkbox(intItemID)
		Response.Write("</tr>")
	End Sub

	Public Sub DisplayStandardDropdownFilter(aryValues, aryNames, intItemID)
		Dim i
		Response.Write("<tr>")
		Call DisplayStandardFilter_Title(intItemID)

		Response.Write("<td valign=""top"">")
		Response.Write("<select name=""" & GetLinkChar(intItemID) & """>")
		For i=0 To UBound(aryNames)
			Call AddSelectOption(aryValues(i), aryValues(i) & " " & aryNames(i), m_aryFilter(intItemID))
		Next
		Response.Write("</select>")
		Response.Write("</td>")

		Call DisplayStandardFilter_Checkbox(intItemID)
		Response.Write("</tr>")
	End Sub

	Public Sub DisplayYesNoFilter(intItemID)
		Response.Write("<tr>")
		Call DisplayStandardFilter_Title(intItemID)

		Response.Write("<td valign=""top"">")
		Response.Write("<select name=""" & GetLinkChar(intItemID) & """>")
		Call AddSelectOption("0", "No", m_aryFilter(intItemID))
		Call AddSelectOption("1", "Yes", m_aryFilter(intItemID))
		Response.Write("</select>")
		Response.Write("</td>")

		Call DisplayStandardFilter_Checkbox(intItemID)
		Response.Write("</tr>")
	End Sub

	Public Sub DisplayStandardFilter_Title(intItemID)
		Response.Write("<td valign=""center"">" & FONT_SMALL & Replace(m_aryTitle(intItemID), " ", "&nbsp;") & "&nbsp;</font></td>")
	End Sub

	Public Sub DisplayStandardFilter_Control(intItemID)
		Response.Write("<td valign=""top""><input type=""text"" name=""" & GetLinkChar(intItemID) & """")
		If m_aryDataType(intItemID) And tDATE Then
			Response.Write(" value=""" & NiceDate(m_aryFilter(intItemID)))
		Else
			Response.Write(" value=""" & m_aryFilter(intItemID))
		End If		
		Response.Write(""">&nbsp;</td>")
	End Sub

	Public Sub DisplayStandardFilter_Checkbox(intItemID)
		Response.Write("<td valign=""center""><input type=""checkbox"" name=""B" & GetLinkChar(intItemID) & """ ")
		If m_aryInclude(intItemID) Then Response.Write("checked")
		Response.Write(">&nbsp;</td>")
	End Sub

	' --------------------------------------------------------------------
	Public Function GetLinkChar(intEntry)
		'Gets the URL/Form name for this Entry
		GetLinkChar = Right("0" & intEntry, 2)
	End Function

	Public Function GetEntryFromLinkChar(strLinkChar)
		' Finds the entry associated with the given link characters...
		GetEntryFromLinkChar = Int(strLinkChar)
	End Function

	' --------------------------------------------------------------------

	Public Sub GetParameters()
		' Get all parameters from either the querystring or form
		Dim i
		If Request.Form("order") <> "" Or Request.QueryString("order") <> "" Then  
			' We have submitted something at least once, else use the defaults provided
			For i=0 to m_intNumFields-1
				Call GetParameter(i)
				Call GetBooleanParameter(i)
			Next
		End If
	End Sub

	Public Sub GetParameter( intItemID )
		If Request.Form(GetLinkChar(intItemID)) <> "" Then
			' Form overrides querystring
			m_aryFilter(intItemID)	= Trim(Request.Form(GetLinkChar(intItemID)))
			' Better reset the page too
			m_intCurrentPage = CLng(1)
		ElseIf Request.QueryString(GetLinkChar(intItemID)) <> "" Then
			m_aryFilter(intItemID) = Trim(Request.QueryString(GetLinkChar(intItemID)))
		End If
	End Sub
	
	Public Sub GetBooleanParameter( intItemID )
		If Request.Form("B" & GetLinkChar(intItemID)) = "on" Then
			' Form overrides querystring
			m_aryInclude(intItemID)	= True
			' Better reset the page too
			m_intCurrentPage = CLng(1)
		ElseIf Request.QueryString("B" & GetLinkChar(intItemID)) = "on" Then
			m_aryInclude(intItemID)	= True
		Else
			m_aryInclude(intItemID)	= False
		End If
	End Sub

	Public Sub SetPage()
		' Sets the current page if we have specified one.
		If Request("pg") <> "" Then
			m_intCurrentPage = CLng(Request("pg"))
		Else
			m_intCurrentPage = CLng(1)
		End If
	End Sub

	Public Function GetPageControls(intTotalRecords, intTotalPages)
		Dim i
		Dim intFirstPage, intLastPage

		intFirstPage	= 1
		intLastPage		= intFirstPage + 14
		Do While m_intCurrentPage > intLastPage - 6
			intFirstPage	= intFirstPage + 5
			intLastPage		= intLastPage + 5
		Loop
		If intLastPage > intTotalPages Then
			intLastPage = intTotalPages
		End If
		If intLastPage = intTotalPages Then
			Do While intFirstPage > 1 And intFirstPage > intTotalPages -14
				intFirstPage = intFirstPage - 1
			Loop
		End If

		GetPageControls = GetPageControls & "[" & intTotalRecords & "&nbsp;records]&nbsp;"
		
		If Int(m_intCurrentPage) > 1 Then
			GetPageControls = GetPageControls & Link(GetOrder(), Int(m_intCurrentPage)-1) & "&lt;&lt; Previous Page</a>&nbsp;"
		Else
			GetPageControls = GetPageControls & "&lt;&lt; Previous Page&nbsp;"
		End If

		For i=intFirstPage to intLastPage
			If Int(i) = Int(m_intCurrentPage) Then
				GetPageControls = GetPageControls & "<b>" & i & "</b>&nbsp;"
			Else
				GetPageControls = GetPageControls & Link(GetOrder(), i) & i & "</a>&nbsp;"
			End If
		Next

		If Int(m_intCurrentPage) < Int(intTotalPages) Then
			GetPageControls = GetPageControls & Link(GetOrder(), Int(m_intCurrentPage)+1) & "Next Page &gt;&gt;</a>"
		Else
			GetPageControls = GetPageControls & "Next Page &gt;&gt;"
		End If
	
		If intTotalPages > 15 Then
			GetPageControls = GetPageControls & "&nbsp;[" & intTotalPages & "&nbsp;pages]"
		End If

	End Function

	Public Function Link(strOrder, intPageNumber)
		Dim i

		Link = "<a href=""" & Request.ServerVariables("SCRIPT_NAME") & "?order=" & strOrder
		If intPageNumber <> ""		Then Link = Link & "&pg=" & intPageNumber

		For i=0 To m_intNumFields - 1
			If m_aryFilter(i) <> "" Then Link = Link & "&" & GetLinkChar(i) & "=" & Server.URLEncode(m_aryFilter(i))
			' Checkbox booleans...
			If m_aryInclude(i)		Then Link = Link & "&B" & GetLinkChar(i) & "=on"
		Next

		Link = Link & """>"
	End Function

	Public Function GetOrder()
		GetOrder = m_strOrderDirection + GetLinkChar(m_intOrder)
	End Function

	Public Sub SetOrder()
		Dim strOrder
		strOrder = Request("order")

		If Len(strOrder) = 3 Then
			' The ordering has been changed, update accordingly
	
			' strOrder will be a 3 byte string 
			' (1st character = direction)
			' (2nd & 3rd character = sort type)
	
			If Mid(strOrder, 1, 1) = "A" Then
				m_strOrderDirection	= "A"
			Else
				m_strOrderDirection	= "D"
			End If
	
			m_intOrder = GetEntryFromLinkChar( Mid(strOrder, 2, 2) )
		Else
			' Use default
			' TODO - Need to make sure that this isn't a TEXT field
			m_strOrderDirection	= "D"
			m_intOrder			= GetEntryFromLinkChar( Mid(GetSelectedFields(), 1, 2) )
		End If
	End Sub

	Public Sub DisplayColumnHeader(intItemID)
		Dim strUpImg, strDownImg
		strUpImg	= "<img src=""/images/up.gif"" width=""14"" height=""14"" border=""0"">"
		strDownImg	= "<img src=""/images/down.gif"" width=""14"" height=""14"" border=""0"">"

		Response.Write("<td nowrap>" & FONT_STANDARD )

		Dim intTest
		intTest = tID Or tTEXT
		
		If (m_intOrder = intItemID) Then 
			Response.Write("<b>" & m_aryTitle(intItemID) & "</b>") 
		Else
			If Not((m_aryDataType(intItemID) And tTEXT) = tTEXT) Then
				' We can't "ORDER BY" for text fields...
				Response.Write(Link("A" & GetLinkChar(intItemID), m_intCurrentPage)) 
			End If			
			Response.Write(m_aryTitle(intItemID))
			If Not((m_aryDataType(intItemID) And tTEXT) = tTEXT) Then
				Response.Write("</a>")
			End If
		End If

		If (m_intOrder = intItemID) Then 
			If (m_strOrderDirection = "D") Then
				Response.Write("&nbsp;" & Link("A" & GetLinkChar(intItemID), m_intCurrentPage) & strDownImg & "</a>")
			Else
				Response.Write("&nbsp;" & Link("D" & GetLinkChar(intItemID), m_intCurrentPage) & strUpImg & "</a>")
			End If
		End If
		Response.Write("</td>")
	End Sub 

	Public Sub DisplayResultsHeader()
		Dim strLine
		Dim i
		Dim strSelectedFields
		Dim intEntry

		strSelectedFields = GetSelectedFields()

		Response.Write("<tr bgcolor=""#E0E0E0"">")

		For i=1 To Len(strSelectedFields) Step 2
			intEntry = GetEntryFromLinkChar(Mid(strSelectedFields, i, 2))
			Call DisplayColumnHeader(intEntry)
		Next	
		Response.Write("</tr>")
	End Sub

	Public Function GenerateSQL(intFieldStyle)
		Dim strWhereClause
		Dim strSubWhereClause
		Dim strFields
		Dim i
		Dim intEntry
		Dim strSelectedFields

		strSelectedFields = GetSelectedFields()

		' GENERATE Select/Fields/Table portion ------------------------------------------------
		If intFieldStyle = SELECTEDFIELDS Then
			strFields = ""
			For i=1 To Len(strSelectedFields) Step 2
				intEntry = GetEntryFromLinkChar(Mid(strSelectedFields, i, 2))
	
				If strFields <> "" Then strFields = strFields & ", "
				strFields = strFields & m_aryFieldName(intEntry)
			Next
		Else
			strFields = ""
			For i=0 to m_intNumFields-1
				If strFields <> "" Then strFields = strFields & ", "
				strFields = strFields & m_aryFieldName(i)
			Next
		End If

		GenerateSQL = "SELECT " & strFields & " FROM " & m_strTableName & " "

		' Generate WHERE clause ----------------------------------------------------------------
		strWhereClause = ""

		For i=1 To Len(strSelectedFields) Step 2 
			intEntry = GetEntryFromLinkChar(Mid(strSelectedFields, i, 2))

			If m_aryInclude(intEntry) Then  'If this field is part of our filter...
				strSubWhereClause = ""
				Select Case m_aryFilterStyle(intEntry) 
					Case fEQUAL	strSubWhereClause = m_aryFieldName(intEntry) & " = " 
					Case fGT_EQ	strSubWhereClause = m_aryFieldName(intEntry) & " >= "
					Case fLT_EQ	strSubWhereClause = m_aryFieldName(intEntry) & " <= "
					Case fLIKE	strSubWhereClause = m_aryFieldName(intEntry) & " LIKE "
				End Select
		
				If (m_aryDataType(intEntry) And tDATE) Or (m_aryDataType(intEntry) And tTEXT) Or (m_aryDataType(intEntry) And tVARCHAR) Then 
					strSubWhereClause = strSubWhereClause & "'"
				End If
				If m_aryFilterStyle(intEntry) = fLIKE 		Then strSubWhereClause = strSubWhereClause & "%"
		
				' Handle GMN specific data types etc. ----------------------------------------------------
				If (m_aryDataType(intEntry) And tMUZE) Then 
					m_aryFilter(intEntry) = Space(7-Len(m_aryFilter(intEntry))) & m_aryFilter(intEntry)
				End If
				'------------------------------------------------------------------------------------------

				strSubWhereClause = strSubWhereClause & Replace(m_aryFilter(intEntry), "'", "''")
		
				If m_aryFilterStyle(intEntry) = fLIKE 		Then strSubWhereClause = strSubWhereClause & "%"
				If (m_aryDataType(intEntry) And tDATE) Or (m_aryDataType(intEntry) And tTEXT) Or (m_aryDataType(intEntry) And tVARCHAR) Then 
					strSubWhereClause = strSubWhereClause & "'"
				End If

				strWhereClause = strWhereClause & AddWhereClause(strWhereClause, strSubWhereClause)
			End If
		Next

		If strWhereClause <> "" Then
			GenerateSQL = GenerateSQL & " WHERE " & strWhereClause & " "
		End If
		' ----------------------------------------------------------------------------------------
		
		' Generating ORDER BY clause ------------------------------------------------------------
		GenerateSQL = GenerateSQL & "ORDER BY " & m_aryFieldName(m_intOrder)

		If m_strOrderDirection	= "D" Then
			GenerateSQL = GenerateSQL & " DESC"
		End If
		' ---------------------------------------------------------------------------------------
	End Function

	Public Sub DisplayHeader()
		Dim objTheObject

		Select Case m_strAction
			Case HOME, VIEW	
'				objTabbedMenu.m_intSelectedTab = 0
				Set objTheObject = Me
'			Case CONFIGURE	
'				objTabbedMenu.m_intSelectedTab = 1
'				Set objTheObject = New CDBConfigure
'				Call objTheObject.SetParentObject(Me)
'			Case REPORT, SENDREPORT		
'				objTabbedMenu.m_intSelectedTab = 2
'				Set objTheObject = New CDBReport
'				Call objTheObject.SetParentObject(Me)
		End Select
%>	
<html>
<head>
<title><%=m_strApplicationName%></title>
</head>
<% 
		Call Display()

		Select Case m_strAction
			Case CONFIGURE, REPORT, SENDREPORT	Set objTheObject = Nothing
		End Select

'		Set objTabbedMenu = Nothing
	End Sub

	Public Sub Display()
	%>
<body>

<SCRIPT Language="JavaScript">
function showSearch(){
	if (document.getElementById('searchdiv').style.display=='none')
		document.getElementById('searchdiv').style.display='block'
	else
		document.getElementById('searchdiv').style.display='none'
}
</SCRIPT>

<%=FONT_SMALL%><A href="javascript:showSearch()">TOGGLE SEARCH BAR</A><BR>
<% If m_strAction = VIEW Then %>
<DIV ID="searchdiv" style="display:none;">
<% Else %>
<DIV ID="searchdiv" style="display:block;">
<% End If %>


<table bgcolor="#D6D6CE" width="100%" cellpadding="0" cellspacing="0" border="0">
<form action="<%=Request.ServerVariables("URL")%>" name="DBViewer" method="post">
<input type="hidden" name="Action" value="">
<tr>
	<td>
		<table width="100%" cellpadding="2" cellspacing="2" border="0">
		<tr>
			<td width="50%" valign="top">
			<% Call DisplayFilters() %>
			</td>
			<td width="50%" valign="top"><%=FONT_STANDARD%>
				<a style="text-decoration:none" href="javascript:document.DBViewer.Action.value='<%=VIEW%>';javascript:document.DBViewer.submit();">
				View
				</a><br><br>
			</td>
		</tr>
		</table>
	</td>
</tr>
</form>
</table>
</DIV>
	<%
	End Sub

	Public Sub DisplayFilters()
		Dim strSelectedFields
		strSelectedFields = GetSelectedFields()
		Dim intEntry
		Dim i
%>
	<table cellpadding="0" cellspacing="0" border="0">
	<input type="hidden" name="order" value="<%=GetOrder()%>">
<%
		For i=1 To Len(strSelectedFields) Step 2
			intEntry = GetEntryFromLinkChar(Mid(strSelectedFields, i, 2))
			Call m_objParent.DisplayFilter(intEntry)
		Next
%>
	</table>
<%
	End Sub

	Public Sub DisplayEntries()
		Dim sql, intRow, i, strPageControls, intTotalPages, intTotalRecords
		Dim strLink
		Dim intEntry
		Dim strSelectedFields
		strSelectedFields = GetSelectedFields()

		sql = GenerateSQL(SELECTEDFIELDS)

		m_connConnection.Open m_strDSN

		' Handle pages ------------------------------------------------------------------------
		m_rs.PageSize		= m_intPageSize
		m_rs.CursorLocation	= 2
		m_rs.CacheSize		= m_intPageSize * 10

		m_rs.Open sql, m_connConnection, 3, 3
		intTotalRecords = m_rs.RecordCount
		intTotalPages	= m_rs.PageCount
		If m_intCurrentPage > intTotalPages Then
			m_intCurrentPage = intTotalPages 
		End If
		
		If Not m_rs.EOF Then
			m_rs.AbsolutePage = m_intCurrentPage
		End If

		If intTotalPages > 1 Then
			strPageControls = GetPageControls(intTotalRecords, intTotalPages)
			Response.Write("<br><center>" & FONT_SMALL & strPageControls & "</center><p>")
		End If

		' Display a page full of entries  --------------------------------------------------------
%>
		<table cellpadding="2" cellspacing="2" border="0" width="100%">
<%		
		Call DisplayResultsHeader()
		For intRow=1 to m_intPageSize
			If m_rs.EOF Then
				Exit For
			End If

			Call GetRecordSet()

			If m_bHighlight Then
				Response.Write("<tr bgcolor=""#FOFOFO"">")
			Else
				Response.Write("<tr>")
			End If

			For i=1 To Len(strSelectedFields) Step 2
				intEntry = GetEntryFromLinkChar(Mid(strSelectedFields, i, 2))
				Response.Write("<td nowrap valign=""top"">" & FONT_SMALL)
				' See if we need to link to this item -------------------
				strLink = m_objParent.LinkEdit(intEntry)				
				If strLink <> "" Then Response.Write(strLink)
				'--------------------------------------------------------
				If m_aryDataType(intEntry) And tDATE Then
					Response.Write(NiceDate(m_aryData(intEntry)))
				Else
					Response.Write(m_objParent.MapData(intEntry))
				End If

				If strLink <> "" Then Response.Write("</a>")

				Response.Write("</td>")
			Next				
			Response.Write("</tr>")
			
			m_bHighlight = Not m_bHighlight
			m_rs.MoveNext
		Next

		m_rs.Close
		m_connConnection.Close
%>
		</table>
<%
		'----------------------------------------------------------------------------------------
	End Sub

	Public Function MapValue(aryInput, aryOutput, intEntryID)
		Dim i, bFound
		bFound = False
		For i=0 to UBound(aryInput)
			If CLng(aryInput(i)) = CLng(m_aryData(intEntryID)) Then
				bFound = True
				MapValue = aryOutput(i)
				Exit For
			End If
		Next
		If bFound = False Then MapValue = m_aryData(intEntryID)
	End Function

	Public Sub GetRecordSet()
		Dim i
		Dim strSelectedFields
		Dim intEntryID
		strSelectedFields = GetSelectedFields()

		For i=1 To Len(strSelectedFields) Step 2
			intEntryID = GetEntryFromLinkChar(Mid(strSelectedFields, i, 2))
			If m_aryInclude(intEntryID) And m_aryFilter(intEntryID) <> "" Then
				If (m_aryDataType(intEntryID) And tTEXT) Or (m_aryDataType(intEntryID) And tVARCHAR) Then
					Call FormatTextField(intEntryID)
				Else
					m_aryData(intEntryID) = Trim(m_rs(m_aryFieldName(intEntryID)))
				End If
			Else
				m_aryData(intEntryID) = Trim(m_rs(m_aryFieldName(intEntryID)))
			End If
		Next
	End Sub

	Public Sub FormatTextField(intEntryID)
		' TODO - This can probably be optimised big style
		' TODO - Put additional error checking in here to

		' Trims TEXT fields and also highlights any filter words if we have any

		Dim intTextFieldCutOff
		intTextFieldCutOff = 40

		Dim intFoundPos
		Dim strLowerCaseFilter
		Dim strLowerCaseField
		Dim strLeft
		Dim strRight
		Dim strMiddle
		' ----------------------------------------

		m_aryData(intEntryID) = Trim(m_rs(m_aryFieldName(intEntryID)))
		m_aryData(intEntryID) = StripHTML(m_aryData(intEntryID))

		' So we can compare properly (case insensitive like SQL)
		strLowerCaseFilter	= LCase(m_aryFilter(intEntryID))
		strLowerCaseField	= LCase(m_aryData(intEntryID))

		intFoundPos	= Instr(strLowerCaseField, strLowerCaseFilter)
		
		If intFoundPos > 0 Then ' This should happen!
			If Len(m_aryFilter(intEntryID)) < intTextFieldCutOff Then
				' Centre the found text
				intFoundPos = intFoundPos - ( intTextFieldCutOff - Len(Trim(m_aryFilter(intEntryID))) ) / 2
				If intFoundPos < 1 Then	intFoundPos = 1
			End If
			m_aryData(intEntryID) = Mid(m_aryData(intEntryID), intFoundPos, intTextFieldCutOff)
		End If

		' Highlight search keyword
		intFoundPos	= Instr(LCase(m_aryData(intEntryID)), LCase(m_aryFilter(intEntryID)))
		If intFoundPos > 0 Then 
			strLeft		= Mid(m_aryData(intEntryID), 1, intFoundPos-1 )
			strMiddle	= Mid(m_aryData(intEntryID), intFoundPos, Len(m_aryFilter(intEntryID)))
			strRight	= Right(m_aryData(intEntryID), (Len(m_aryData(intEntryID))-intFoundPos+1)-Len(m_aryFilter(intEntryID)))
	
			'm_aryData(intEntryID) = "..." & strLeft & "<b>" & strMiddle & "</b>" & strRight & "..."
			m_aryData(intEntryID) = strLeft & "<b>" & strMiddle & "</b>" & strRight
		End If
	End Sub

	Public Function AddWhereClause( strWhereClause, strClause)
		If strWhereClause <> ""	Then 
			AddWhereClause = " And " & strClause
		Else
			AddWhereClause = strClause
		End If
	End Function

	Public Function NiceDate(dateDate)
		If IsDate(dateDate) Then
			NiceDate = Day(dateDate) & "/" & MonthName(Month(dateDate), True) & "/" & Right(Year(dateDate), 2)
		Else
			NiceDate = dateDate
		End If
	End Function

	Public Function NiceTime(dateDate)
		Dim strHour, strMinute
		If IsDate(dateDate) Then
			strHour = Hour(dateDate)
			If Len(strHour) = 1 Then
				strHour = "0" & strHour
			End If

			strMinute = Minute(dateDate)
			If Len(strMinute) = 1 Then
				strMinute = "0" & strMinute
			End If

			NiceTime = strHour & ":" & strMinute
		Else
			NiceTime = dateDate
		End If
	End Function

	Public Sub DisplayFooter()
		Response.Write("</body></html>")
	End Sub

	Public Function GetSelectedFields()
		If m_strPreferredSelectedFields <> "" Then
			GetSelectedFields = m_strPreferredSelectedFields
		Else
			GetSelectedFields = m_strDefaultSelectedFields
		End If
	End Function

	Public Sub AddSelectOption(strValue, strText, strSelectedValue)
		' Adds an entry into an HTML <SELECTion> box
		Response.Write("<option ")
		If Trim(CStr(strValue)) = Trim(CStr(strSelectedValue)) Then
			Response.Write("selected ")
		End If
		Response.Write("value=""" & strValue & """>" & strText)
	End Sub

	Private Function StripHTML(strBody)
		If InStr(strBody, "<") OR InStr(strBody, ">") Then
			Dim iStart, iEnd, iPos
			iStart 	= Instr(1, strBody, "<")
			iEnd	= Instr(1, strBody, ">")
			Do Until iStart = 0 OR iEnd = 0
				If iStart > iEnd Then
					iStart	= iEnd + 1
				Else
					strBody = Left(strBody, iStart-1) & Right(strBody, Len(strBody) - iEnd)
				End If
				iPos	= iStart			'setting string cursor pos
				If iPos = 0 Then iPos = 1
				iStart 	= Instr(iPos, strBody, "<")
				iEnd	= Instr(iPos, strBody, ">")
			Loop
		End If
		StripHTML = strBody
	End Function

End Class
%>