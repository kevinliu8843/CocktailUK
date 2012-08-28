<% 
' CStack is below also...

'TODO - When satisfied, remove debug stuff...

Const URLID = 0
Const WORD	= 1
Const COUNT	= 2

Const TITLES	= 1
Const TEXT		= 2

Const ASCENDING		= 1
Const DESCENDING	= 2
Const TITLE			= 4
Const HITS			= 8

CONST SEARCH_TYPE_GENERAL= 1
CONST SEARCH_TYPE_FORUMS= 2
CONST SEARCH_TYPE_DRINK	= 3
CONST SEARCH_TYPE_SHOP	= 4


Class CSearch
	' Results ------------------------------------
	Public m_strTitle
	Public m_strURL
	Public m_intTypeID
	Public m_fltHitScore
	Public m_intLineNumber
	' ------------------------------------------------

	Public  m_objConn
	Private m_objStopWord
	Private m_aryRecordSet		' Array of URLIDs/WordID/Count 
	Private m_intNumRecords		' Number of records of above
	Private m_objrsResults

	Private m_objWordsDic
	Public  m_intSearchType
	Public	m_strSearchString
	Public	m_intSortOrder
	Private m_intNumTokens

	Private m_objProcessingStack
	Private m_objOperatorStack

	Private m_strErrorText
	Private m_intCharErrorPos		' Low-level error position...
	Private m_intTokenErrorPos		' High-level error position...
	
	Private m_bHasResults


	Public m_intPageSize
	Public m_intCurrentPage
	Public m_intTotalPages
	Private m_intCurrentLine


	Private m_intFilter

	Public m_bShowDebug

	Public Sub Class_Initialize
		Set m_objProcessingStack	= New CStack
		Set m_objOperatorStack		= New CStack
		Set m_objWordsDic			= Server.CreateObject("Scripting.Dictionary")
		Set m_objConn				= Server.CreateObject("ADODB.Connection")
		Set m_objrsResults			= Server.Createobject("ADODB.Recordset")

		m_objConn.Open(strDB)

		m_strErrorText	= ""
		m_intTokenErrorPos	= -1
		m_intCharErrorPos	= -1

		m_intNumTokens	= 0
		m_intNumRecords	= 0

		m_intSearchType = TITLES Or TEXT
		m_intSortOrder	= HITS Or DESCENDING

		m_bHasResults	= False
		m_bShowDebug	= False

		m_intPageSize		= 10
		m_intCurrentPage	= 1
		m_intTotalPages		= 1
	
		Set m_objStopWord = New CStopWord
	End Sub

	Public Sub Class_Terminate
		Set m_objWordsDic			= Nothing
		Set m_objConn				= Nothing
		Set m_objProcessingStack	= Nothing
		Set m_objOperatorStack		= Nothing
		Set m_objStopWord			= Nothing
		Set m_objrsResults			= Nothing
	End Sub

	Public Function HasResults()
		HasResults = m_bHasResults
	End Function

	Public Function GetNextToken(strText, intPos, bPreserve)
		Dim intLen, strChar

		GetNextToken = ""
		intLen = Len(strText)
		
		' Find the start of the next token...
		Do While True
			If intPos > intLen Then Exit Do

			strChar = LCase(Mid(strText, intPos, 1))

			If (strChar >= "a" And strChar <= "z") Or strChar = "(" Or strChar = ")" Then
				Exit Do
			ElseIf strChar <> " " Then
				' Error... (invalid char)
				Call SetError("Syntax Error", intPos, -1)
				intPos = 0
				Exit Function
			End If
			intPos = intPos+1
		Loop

		If intPos > intLen Then 
			intPos = 0			' No more tokens
			Exit Function
		Else
			'Detemine token type from initial character...
			If strChar = "(" Or strChar = ")" Then
				GetNextToken = strChar
				intPos = intPos + 1
			Else
				Do While True
					If intPos > intLen Then Exit Do
					
					strChar = LCase(Mid(strText, intPos, 1))

					If (strChar >= "a" And strChar <= "z") Or strChar = "'" Or strChar = "-" Then
						If bPreserve Then
							GetNextToken = GetNextToken & Mid(strText, intPos, 1)
						Else
							GetNextToken = GetNextToken & strChar
						End If
					Else
						Exit Do
					End If

					intPos = intPos+1
				Loop
				
				If Not bPreserve Then
					If Right(GetNextToken, 2) = "'s" then
						GetNextToken = Left(GetNextToken, Len(GetNextToken)-2)
					End If

					If Len(GetNextToken) > 30 Then
						GetNextToken = Left(GetNextToken, 30)
					End If
				End If

				If GetNextToken = "and" Or GetNextToken = "or" Then
					GetNextToken = UCase(GetNextToken)
				End If
			End If
		End If
	End Function

	Public Sub DoSearch()
		Dim  i, intNumWords

		m_strSearchString = StripHTML(m_strSearchString)
		m_strSearchString = StripSyntaxErrors(m_strSearchString)
		intNumWords = AddWordsToDictionary() 

		If intNumWords > 0 And m_intCharErrorPos=-1 And m_intTokenErrorPos=-1 Then
			' Parse search string, add WORDs to dictionary

			Call GetRecordSet()

			' Based on WORDs generate SQL and get records
			Call AddURLIDsToWordDictionary()	' For each word, generate a URLID list

			Call Evaluate()

			If m_strErrorText = "" Then
				If m_objProcessingStack.Count() = 1 Then
					If m_bShowDebug Then
						Response.Write("<BR>Result = [" & m_objProcessingStack.Top()(0) & "]<BR>")
					End If
					Call GetSuccessfulURLs(m_objProcessingStack.Pop()(0))
				Else 
					Call SetError("Syntax Error", -1, -1)
				End If
			End If

	' ---------------------------------------------------------------
If m_bShowDebug Then
			For i=0 to m_intNumRecords-1
				Response.Write(m_aryRecordSet(URLID, i))
				Response.Write(" - " & m_aryRecordSet(WORD, i))
				Response.Write(" - " & m_aryRecordSet(COUNT, i))

				Response.Write("<BR>")
			Next
			Response.Write("<BR>")

			Dim aryWords
			aryWords = m_objWordsDic.Keys
			i=0
			Do While i<=UBound(aryWords)
				Response.Write(aryWords(i) & " - " & m_objWordsDic.Item(aryWords(i)) )
				Response.Write("<BR>")

				i=i+1
			Loop
End If
		ElseIf intNumWords > 0 Then
			If m_intNumTokens = 0 Then
				Call SetError("No Search String", -1, -1)
			Else
				Call SetError("Syntax Error", -1, 0)
			End If
		End If

		' ---------------------------------------------------------------
	End Sub

	Private Sub GetSuccessfulURLs(strURLList)
		Dim sql, intPos, intHits, intHitScore, strSort, aryFields, strTypeIDs
		aryFields = Array("URL", "Title", "TypeID", "HitScore")

		If strURLList <> "" Then
			m_objrsResults.PageSize			= m_intPageSize
			m_objrsResults.CursorLocation	= 2
			m_objrsResults.CacheSize		= m_intPageSize * 10

			m_objrsResults.Fields.Append "URL",  200, 255	'adVarChar
			m_objrsResults.Fields.Append "Title", 200, 255	'adVarChar
			m_objrsResults.Fields.Append "TypeID", 3			'adInteger
			m_objrsResults.Fields.Append "HitScore", 3		'adInteger
			m_objrsResults.Open

			sql = "SELECT ID, TypeID, URL, Title, "
			If m_intSearchType = TITLES Then
				sql = sql & "TotalTitleWords AS Total "
			ElseIf m_intSearchType = TEXT Then
				sql = sql & "TotalCopyWords AS Total "
			Else
				sql = sql & "TotalTitleWords+TotalCopyWords AS Total "
			End If
			
			sql = sql & "FROM URLs "
			sql = sql & "WHERE ID IN (" & strURLList & ") "
			strTypeIDs = GetTypeIDs()
			If strTypeIDs <> "" Then
				sql = sql & "AND TypeID IN (" & strTypeIDs & ") "
			End If
			sql = sql & "ORDER BY ID"
			rs.Open sql, m_objConn, 0, 3

			intPos=0
			Do While Not rs.EOF
				intHits = 0
				Do While True
					If intPos > m_intNumRecords-1 Then Exit Do
					
					If Int(m_aryRecordSet(URLID, intPos)) < Int(rs("ID")) Then
						intPos = intPos + 1
					Else 
						Exit Do
					End If
				Loop

				If intPos > m_intNumRecords-1 Then Exit Do

				If Int(rs("ID")) = Int(m_aryRecordSet(URLID, intPos)) Then
					Do While True
						If intPos > m_intNumRecords-1 Then Exit Do
						
						If Int(rs("ID")) = Int(m_aryRecordSet(URLID, intPos)) Then 
							intHits = intHits + Int(m_aryRecordSet(COUNT, intPos))
							intPos = intPos + 1
						Else 
							Exit Do
						End If
					Loop
				End If

				intHitScore = (CDbl(intHits) / Max(CDbl(rs("Total")),1) ) * 10000
				m_objrsResults.AddNew aryFields, Array(rs("URL"), rs("Title"), rs("TypeID"), intHitScore)

				rs.MoveNext
			Loop
			rs.Close

			If Not m_objrsResults.EOF Then
				If (objSearch.m_intSortOrder And TITLE) Then
					strSort = "Title"
				Else
					strSort = "HitScore"
				End If

				If (objSearch.m_intSortOrder And DESCENDING) Then
					strSort = strSort & " DESC"
				End If 
				m_objrsResults.Sort			= strSort
				m_objrsResults.AbsolutePage	= m_intCurrentPage

				m_bHasResults		= True
				m_intTotalPages		= m_objrsResults.PageCount
				m_intCurrentLine	= 1
			Else
				Call SetError("Your search yielded no results", -1, -1)
			End If
		Else
			Call SetError("Your search yielded no results", -1, -1)
		End If
	End Sub

	Public Function GetFirst()
		Dim aryScale, i, fltHitScore, bFound
		aryScale = Array(2500, 1667, 1111, 741, 494, 329, 219, 146, 98) 
		
		If Not m_objrsResults.EOF Then
			m_strTitle		= m_objrsResults("Title")
			m_strURL		= m_objrsResults("URL")
			m_intTypeID		= Int(m_objrsResults("TypeID"))
			m_fltHitScore	= Int(m_objrsResults("HitScore"))
			m_intLineNumber = m_intCurrentLine + ((m_intCurrentPage-1)*m_intPageSize)

			' Map Hit Score to "Star" number (0 - 5) --------
			fltHitScore = 5.0
			bFound = False
			For i=0 To UBound(aryScale)
				If m_fltHitScore >= aryScale(i) Then
					m_fltHitScore = fltHitScore
					bFound=True
					Exit For
				End If
				fltHitScore = fltHitScore - 0.5
			Next
			If Not bFound Then
				m_fltHitScore = 0.5
			End If
			' -----------------------------------------------

			Select Case m_intTypeID
				Case SEARCH_TYPE_GENERAL 
					m_strURL = "/" & m_strURL
				Case SEARCH_TYPE_FORUMS
					m_strURL = "/forums/messages.asp?id=" & m_strURL
				Case SEARCH_TYPE_DRINK
					m_strURL = "/db/viewCocktail.asp?ID=" & m_strURL
				Case SEARCH_TYPE_SHOP
					m_strURL = "/shop/products/product.asp?ID=" & m_strURL
			End Select

			GetFirst = True
		Else
			GetFirst = False
		End If
	End Function

	Public Function GetNext()
		m_objrsResults.MoveNext
		m_intCurrentLine = m_intCurrentLine + 1
		If m_intCurrentLine > m_intPageSize Then
			GetNext = False
		Else	
			GetNext = GetFirst()
		End If
	End Function

	Private Sub Evaluate()
		Dim intWordPos, strToken, strPreviousToken, intWordNum, aryEntry
		Dim bFound

		m_objOperatorStack.SetSize(m_intNumTokens)	' A good initial size
		m_objOperatorStack.SetSize(m_intNumTokens)	' A good initial size

		intWordPos	= 1
		intWordNum	= 0
		strPreviousToken = ""
		Do While intWordPos >0
			strToken = GetNextToken(m_strSearchString, intWordPos, False)
			If intWordPos > 0 Then

				' Handle IMPLIED ANDs ----------------------------------------------------
				If strPreviousToken <> "" Then
					If strToken = "(" And strPreviousToken = ")" Then
						Call HandleOperatorToken("AND", intWordNum)
					ElseIf IsWord(strToken) And IsWord(strPreviousToken) Then
						Call HandleOperatorToken("AND", intWordNum)
					ElseIf IsWord(strToken) And strPreviousToken =")" Then
						Call HandleOperatorToken("AND", intWordNum)
					End If
				End If

				' Check for errors --------------------------------------------------------
				If strPreviousToken <> "" Then
					If IsOperator(strPreviousToken) And IsOperator(strToken) Then
						Call SetError("Syntax Error", -1, intWordNum)
						Exit Do
					ElseIf strPreviousToken = "(" And IsOperator(strToken) Then
						Call SetError("Syntax Error", -1, intWordNum)
						Exit Do
					ElseIf IsOperator(strPreviousToken) And strToken = ")" Then
						Call SetError("Syntax Error", -1, intWordNum)
						Exit Do
					End If
				End If
				
				If IsWord(strToken) Then
					If m_bShowDebug Then
						Response.Write("[" & strToken & "] ")
					End If
					m_objProcessingStack.Push(Array(m_objWordsDic.Item(strToken), intWordNum))
				ElseIf IsOperator(strToken) Then
					Call HandleOperatorToken(strToken, intWordNum)
				ElseIf strToken = "(" Then
					m_objOperatorStack.Push(Array(strToken, intWordNum))
				ElseIf strToken = ")" Then
					' Search for corresponding "("
					bFound = False
					Do While Not m_objOperatorStack.IsEmpty()
						aryEntry = m_objOperatorStack.Pop()
						If aryEntry(0) = "(" Then
							bFound = True
							Exit Do
						Else
							Call ProcessOperator(aryEntry)
						End If
					Loop
					If bFound = False Then
						Call SetError("Unexpected )", -1, intWordNum)
						Exit Do
					End If
				End If

				strPreviousToken = strToken
				intWordNum = intWordNum + 1
			End If
		Loop

		' Empty out operator stack
		If m_strErrorText = "" Then
			Do While Not m_objOperatorStack.IsEmpty()
				aryEntry = m_objOperatorStack.Pop()

				If aryEntry(0) = "(" Then
					Call SetError("Missing )", -1, aryEntry(1))
					Exit Do
				End If

				Call ProcessOperator(aryEntry)
			Loop
		End If
	End Sub
	
	Private Sub SetError(strText, intErrorCharPos, intErrorTokenPos)
		If m_strErrorText = "" Then
			m_strErrorText	= strText
			If intErrorCharPos >-1 Then m_intCharErrorPos = intErrorCharPos
			If intErrorTokenPos > -1 Then m_intTokenErrorPos = intErrorTokenPos
		End If
	End Sub

	Public Sub ShowError()
		Dim intWordPos, intWordNum, strToken
		intWordPos	= 1
		intWordNum	= 0
		Response.Write(m_strErrorText & "<BR>")

		If m_intCharErrorPos > -1 Then
			Response.Write(Left(m_strSearchString, m_intCharErrorPos-1))
			Response.Write("<font color=""red"">" & Mid(m_strSearchString, m_intCharErrorPos, 1) & "</font>")
			Response.Write(Right(m_strSearchString, (Len(m_strSearchString)-(m_intCharErrorPos-1))-1  ))
		Else
			Do While intWordPos >0
				strToken = GetNextToken(m_strSearchString, intWordPos, True)
				If intWordPos > 0 Then
					If m_intTokenErrorPos = intWordNum Then Response.Write("<b><font color=""red"">")
					Response.Write(strToken & " ")
					If m_intTokenErrorPos = intWordNum Then Response.Write("</b></font>")
				End If
				intWordNum = intWordNum + 1
			Loop
		End If
		Response.Write("<br>")
	End Sub

	Private Sub HandleOperatorToken(strToken, intWordNum)
		Dim aryEntry
		If m_objOperatorStack.IsEmpty() Then
			m_objOperatorStack.Push(Array(strToken, intWordNum))
		ElseIf m_objOperatorStack.Top()(0) = "(" Then
			m_objOperatorStack.Push(Array(strToken, intWordNum))
		Else
			aryEntry = m_objOperatorStack.Top()
			If strToken = "AND" And aryEntry(0) = "OR" Then
				' Only case where current can be higher priority than TOS
				m_objOperatorStack.Push(Array(strToken, intWordNum))
			Else
				Call ProcessOperator(m_objOperatorStack.Pop())
				m_objOperatorStack.Push(Array(strToken, intWordNum))
			End If
		End If
	End Sub

	Private Sub ProcessOperator(aryEntry)
		Dim strLeft, strRight
		'aryEntry(0) = Operator
		'aryEntry(1) = Token Position

		If m_objProcessingStack.Count() >=2 Then
			strRight	= m_objProcessingStack.Pop()(0)
			strLeft		= m_objProcessingStack.Pop()(0)

			If m_bShowDebug Then
				Response.Write("[" & aryEntry(0) & "] ")
			End If

			If aryEntry(0) = "AND" Then
				Call m_objProcessingStack.Push(Array(ANDList(strLeft, strRight), aryEntry(1)))
			ElseIf aryEntry(0) = "OR" Then
				Call m_objProcessingStack.Push(Array(ORList(strLeft, strRight), aryEntry(1)))
			Else
				Call SetError("Syntax Error", -1, aryEntry(1))
			End If
		Else
			Call SetError("Syntax Error", -1, aryEntry(1))
		End If
	End Sub

	Private Function GetRecordSet()
		Dim aryWords
		Dim sql, i

		If m_objWordsDic.Count > 0 Then
			sql = "SELECT WordLink.URLID, Words.Word, "

			If m_intSearchType = TITLES Then
				sql = sql & "WordLink.TitleCount "
			ElseIf m_intSearchType = TEXT Then
				sql = sql & "WordLink.CopyCount "
			Else
				sql = sql & "(WordLink.TitleCount*7) + WordLink.CopyCount "
			End If

			sql = sql & "FROM Words INNER JOIN WordLink ON Words.ID = WordLink.WordID "
			sql = sql & "WHERE ("
			
			aryWords = m_objWordsDic.Keys
			i=0
			Do While i<=UBound(aryWords)
				If i>0 Then sql = sql & " Or "
				sql = sql & "Words.Word = '" & Replace(aryWords(i), "'", "''") & "' "
				i=i+1
			Loop
			
			sql = sql & ") "

			If m_intSearchType = TITLES Then
				sql = sql & " AND (WordLink.TitleCount>0) "
			ElseIf m_intSearchType = TEXT Then
				sql = sql & " AND (WordLink.CopyCount>0) "
			End If

			sql = sql & " ORDER BY WordLink.URLID"
			rs.Open sql, m_objConn, 0, 3

			If NOT rs.EOF Then
				m_aryRecordSet = rs.GetRows
				m_intNumRecords = UBound(m_aryRecordSet, 2) + 1
				GetRecordSet = True
			Else
				GetRecordSet = False
				m_intNumRecords = 0
			End If

			rs.Close
		Else
			GetRecordSet = False
			m_intNumRecords = 0
		End If
	End Function

	Private Sub AddURLIDsToWordDictionary()
		Dim i
		If m_intNumRecords > 0 Then
			' Add URL IDs for each word...
			For i=0 to m_intNumRecords-1
				If m_objWordsDic.Item(m_aryRecordSet(WORD, i)) <> "" Then 
					m_objWordsDic.Item(m_aryRecordSet(WORD, i)) = m_objWordsDic.Item(m_aryRecordSet(WORD, i)) & ","
				End If
				m_objWordsDic.Item(m_aryRecordSet(WORD, i)) = m_objWordsDic.Item(m_aryRecordSet(WORD, i)) & m_aryRecordSet(URLID, i)
			Next
		End If
	End Sub

	Private Function AddWordsToDictionary()
		' Parse search string and add a dictionary entry for each WORD.  Ready to store a URL List...
		' returns the number of (DISTINCT) words added...
		Dim intWordPos, intWordNum, strToken
		AddWordsToDictionary	= 0
		m_intNumTokens			= 0
		intWordPos = 1
		intWordNum = 0
		Do While intWordPos >0
			strToken = GetNextToken(m_strSearchString, intWordPos, False)
			If intWordPos > 0 Then
				m_intNumTokens = m_intNumTokens + 1
				If IsWord(strToken) Then
					If m_objStopWord.IsStopWord(strToken) Then
						Call SetError("This word has not been indexed", -1, intWordNum)
						Exit Function
					Else
						If Not m_objWordsDic.Exists(strToken) Then
							m_objWordsDic.Add strToken, ""
							AddWordsToDictionary = AddWordsToDictionary + 1
						End If
					End If
				End If
			End If
		intWordNum = intWordNum + 1
		Loop
	End Function

	Private Function IsOperator(strToken)
		IsOperator = (strToken = "OR" Or strToken = "AND")
	End Function

	Private Function IsWord(strToken)
		IsWord = (strToken <> "(" And strToken <> ")" And Not IsOperator(strToken) )
	End Function

	Private Function ANDList(strLeft, strRight)
		' strLeft	= 1,2,3,  strRight	= 2,4,5
		' OUT		= 2
		' strLeft and strRight Must be in sorted Order

		Dim aryLeft, aryRight, i, j

		If strLeft = "" Or strRight = "" Then
			ANDList = ""
		Else 
			aryLeft		= Split(strLeft, ",")
			aryRight	= Split(strRight, ",")

			i=0
			j=0

			Do While True
				If Int(aryLeft(i)) = Int(aryRight(j)) Then
					If ANDList <> "" Then ANDList = ANDList & ","
					ANDList = ANDList & aryLeft(i)
					i=i+1
					j=j+1
				ElseIf Int(aryLeft(i)) < Int(aryRight(j)) Then
					i=i+1
				Else
					j=j+1
				End If

				If i > UBound(aryLeft) Then Exit Do
				If j > UBound(aryRight) Then Exit Do
			Loop
		End If
	End Function

	Private Function ORList(strLeft, strRight)
		' strLeft	= 1,2,3, strRight	= 2,4,5
		' OUT		= 1,2,3,4,5
		' strLeft and strRight Must be in sorted Order

		Dim aryLeft, aryRight, i, j

		If strLeft <> "" And strRight <> "" Then
			aryLeft		= Split(strLeft, ",")
			aryRight	= Split(strRight, ",")

			i=0
			j=0

			Do While True
				If i<=UBound(aryLeft) And j<=UBound(aryRight) Then
					If Int(aryLeft(i)) = Int(aryRight(j)) Then
						If ORList <> "" Then ORList = ORList & ","
						ORList = ORList & aryLeft(i)
						i=i+1
						j=j+1
					ElseIf Int(aryLeft(i)) < Int(aryRight(j)) Then
						If ORList <> "" Then ORList = ORList & ","
						ORList = ORList & aryLeft(i)
						i=i+1
					Else
						If ORList <> "" Then ORList = ORList & ","
						ORList = ORList & aryRight(j)
						j=j+1
					End If
				ElseIf i<=UBound(aryLeft) Then
					If ORList <> "" Then ORList = ORList & ","
					ORList = ORList & aryLeft(i)
					i=i+1
				Else
					If ORList <> "" Then ORList = ORList & ","
					ORList = ORList & aryRight(j)
					j=j+1
				End If
					
				If i > UBound(aryLeft) And j > UBound(aryRight) Then Exit Do
			Loop
		ElseIf strLeft <> "" Then
			ORList = strLeft
		ElseIf strRight <> "" Then
			ORList = strRight
		Else
			ORList = ""
		End If
	End Function

	Public Function StripHTML(strHTML)
		'Strips the HTML tags from strHTML
		
		Dim objRegExp, strOutput
		Set objRegExp = New Regexp
		
		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		objRegExp.Pattern = "<[^>]*>"
		
		'Replace all HTML tag matches with the empty string
		strOutput = objRegExp.Replace(strHTML, "")
		
		stripHTML = strOutput    'Return the value of strOutput
		
		Set objRegExp = Nothing
	End Function
		
	Function StripSyntaxErrors(strText)
		Dim strChar, strOut, i
		strOut = ""
		For i=1 To Len(strText)
			strChar = LCase(Mid(strText, i, 1))
			If ((strChar >= "a" And strChar <= "z") Or strChar = "(" Or strChar = ")" Or strChar = " ") Then
				strOut = strOut & strChar
			End If
		Next
		StripSyntaxErrors = strOut
	End Function

	Public Sub AddFilter(intFilter)
		m_intFilter = m_intFilter Or (2 ^ intFilter)
	End Sub

	Private Function GetTypeIDs()
		Dim i
		GetTypeIDs = ""
		For i= SEARCH_TYPE_GENERAL To SEARCH_TYPE_SHOP
			If m_intFilter And (2 ^ i) Then 
				If GetTypeIDs <> "" Then GetTypeIDs = GetTypeIDs & ","
				GetTypeIDs = GetTypeIDs & i
			End If
		Next
	End Function
End Class

Class CStack
	Private m_aryStack
	Private m_intTop
	Public Sub Class_Initialize
		Redim m_aryStack(0)
		m_intTop = -1
	End Sub

	Public Sub SetSize(intSize)
		Redim m_aryStack(intSize)
	End Sub

	Public Function Count()
		Count = m_intTop+1
	End Function

	Public Sub Push(strValue)
		m_intTop = m_intTop + 1
		If m_intTop > UBound(m_aryStack) Then
			Redim Preserve m_aryStack(m_intTop)
		End If
		m_aryStack(m_intTop) = strValue
	End Sub

	Public Function Pop()
		If m_intTop > -1 Then
			Pop = m_aryStack(m_intTop)
			m_intTop = m_intTop - 1
		Else
			Pop = ""
		End If
	End Function

	Public Function Top()
		Top = m_aryStack(m_intTop)
	End Function

	Public Function IsEmpty()
		IsEmpty = (m_intTop = -1)
	End Function
End Class
%>