<%
Const TITLE_CHUNK	= "T"  
Const COPY_CHUNK	= "C"

' These exist in csearch.inc too, keep in sync.
CONST SEARCH_TYPE_GENERAL= 1
CONST SEARCH_TYPE_FORUMS= 2
CONST SEARCH_TYPE_DRINK	= 3
CONST SEARCH_TYPE_SHOP	= 4


Class CSearchParser
	Private m_objConn
	Private m_objStopWord
	Public	m_strUsername
	Public	m_strPassword

	Public	m_strBaseHREF
	Private m_intNumWordsAdded
	Private m_rs

	Private m_objWordsDictionary

	Public Sub Class_Initialize
		Set m_objConn	= Server.CreateObject("ADODB.Connection")
		Set m_rs		= Server.CreateObject("ADODB.RecordSet")
		m_objConn.Open(strDBMod)

		Set m_objStopWord = New CStopWord
		m_strUsername = ""
		m_strPassword = ""
		Set m_objWordsDictionary = CreateObject("Scripting.Dictionary")
	End Sub

	Public Sub Class_Terminate
		Set m_objConn				= Nothing
		Set m_rs					= Nothing
		Set m_objStopWord			= Nothing
		Set m_objWordsDictionary	= Nothing
	End Sub
	
	Public Sub ParseURL(strURL)
		' Handle a tokenised .asp page...
		Dim strHTML, intURLID, strText, strTitle, intChunkPos, intWordPos, sql
		Dim objXMLHTTP

		Set objXMLHTTP = Server.CreateObject("Msxml2.ServerXMLHTTP.4.0")

		m_intNumWordsAdded = 0

		objXMLHTTP.Open "GET", m_strBaseHREF & strURL, False, m_strUsername, m_strPassword
		objXMLHTTP.Send

		Response.Write(strURL & "<BR>")

		strHTML = objXMLHTTP.ResponseText

		' Handle Title CHUNK ----------------------------------------------------------------
		strTitle = ""
		intChunkPos=1
		strTitle = GetTitle(strHTML)
		Response.Write(strTitle & "<BR>")

		If intChunkPos > 0 Then
			strTitle = Trim(StripHTML(strTitle))
		End If
		intURLID = AddURL(SEARCH_TYPE_GENERAL, RemoveDomain(strURL), strTitle)
		Call ResetURL(intURLID) 
		If strTitle <> "" Then
			Call AddAllWords(TITLE_CHUNK, intURLID, strTitle)
		End If

		' Handle Copy CHUNKs ----------------------------------------------------------------

		intChunkPos=1
		Do While intChunkPos>0
			strText = GetNextChunk(COPY_CHUNK, strHTML, intChunkPos)

			If intChunkPos > 0 Then
				Call AddAllWords(COPY_CHUNK, intURLID, strText)
			End If
		Loop

		Call UpdateTotals(intURLID)
		Response.Write("<BR>")

		Set objXMLHTTP = Nothing
	End Sub

	Public Sub ParseText(intURLType, strForumMessageID, strTitle, strText)
		Dim intURLID, sql

		m_intNumWordsAdded = 0

		intURLID = AddURL(intURLType, strForumMessageID, strTitle)
		Call ResetURL(intURLID) 
		Call AddAllWords(TITLE_CHUNK, intURLID, strTitle)
		Call AddAllWords(COPY_CHUNK, intURLID, strText)
		Call UpdateTotals(intURLID)
	End Sub

	Private Sub UpdateTotals(intURLID)
		Dim sql, intTitleCount, intCopyCount

		' Get Totals ------------------------------------------------
		sql = "SELECT SUM(TitleCount) AS TitleSum, SUM(CopyCount) AS CopySum FROM WordLink "
		sql = sql & "WHERE URLID = " & intURLID
		m_rs.Open sql, m_objConn, 0, 3
		If Not m_rs.EOF Then
			intTitleCount	= m_rs("TitleSum")
			intCopyCount	= m_rs("CopySum")
		End If
		m_rs.Close

		' Update totals and date last indexed ------------------------
		sql = "UPDATE URLs SET TotalTitleWords = " & intTitleCount & ", "
		sql = sql & "TotalCopyWords = " & intCopyCount & " "
		'sql = sql & "LastIndexed = '" & Now()
		sql = sql & " WHERE ID = " & intURLID
		'response.write sql
		m_objConn.Execute(sql)
	End Sub

	Public Sub ClearURL(strURL, intTypeID)
		' URL = Either URL, Forum ID etc.
		Dim intURLID, sql
		intURLID = GetURLID(strURL, intTypeID)
		If intURLID > -1 Then
			Call ResetURL(intURLID) ' Clear off WordLink entries for this ID...

			sql = "UPDATE URLs SET "
			sql = sql & "Deleted = 1, "
			sql = sql & "TypeID = 0, "
			sql = sql & "TotalTitleWords = 0, "
			sql = sql & "Title = 'DELETED', "
			sql = sql & "URL = '-', "
			sql = sql & "LastIndexed = GETDATE()"
			sql = sql & "WHERE ID = " & intURLID

			m_objConn.Execute(sql)
		End If
	End Sub

	Public Sub AddAllWords(strChunkType, intURLID, strText)

		' For this Text Chunk and URL add all words to the various tables...
		Dim intWordPos, intWordID, strWord

		strText = StripHTML(strText)
		intWordPos = 1
		Do While intWordPos >0
			strWord = GetNextWord(strText, intWordPos)
			If intWordPos > 0 And Not m_objStopWord.IsStopWord(strWord) Then
				m_intNumWordsAdded = m_intNumWordsAdded + 1
				intWordID = AddWord(strWord)
				Call AddWordLink(strChunkType, intWordID, intURLID)
			End If
		Loop
	End Sub

	Public Sub AddWordLink(strChunkType, intWordID, intURLID)
		Dim sql

		sql = "SELECT TOP 1 WordID FROM WordLink WHERE WordID = " & intWordID & " AND URLID = " & intURLID
		m_rs.Open sql, m_objConn, 0, 3
		If NOT m_rs.EOF Then
			Select Case strChunkType
				Case TITLE_CHUNK
					sql = "UPDATE WordLink SET TitleCount=TitleCount+1 WHERE WordID = " & intWordID & " AND URLID = " & intURLID
				Case COPY_CHUNK
					sql = "UPDATE WordLink SET CopyCount=CopyCount+1 WHERE WordID = " & intWordID & " AND URLID = " & intURLID
				Case Else
					sql = ""
			End Select
		Else
			sql = "INSERT INTO WordLink (WordID, URLID, TitleCount, CopyCount) VALUES ("
			sql = sql & intWordID & ", "
			sql = sql & intURLID & ", "
			If strChunkType = TITLE_CHUNK Then
				sql = sql & "1, "
			Else
				sql = sql & "0, "
			End If

			If strChunkType = COPY_CHUNK Then
				sql = sql & "1)"
			Else
				sql = sql & "0)"
			End If
		End If
		m_rs.Close
		m_objConn.Execute(sql)
	End Sub

	Public Function AddWord(strWord)
		' Add this word to the db if it's not there, return the WordID
		Dim sql

		If m_objWordsDictionary.Exists(strWord) Then
			AddWord = m_objWordsDictionary.Item(strWord)
		Else
			sql = "SELECT TOP 1 * FROM Words WHERE Word = '" & Replace(strWord, "'", "''") & "'"
			m_rs.Open sql, m_objConn, 0, 3

			If m_rs.EOF Then
				m_rs.Close
				'Add the word
				sql = "INSERT INTO Words (Word) VALUES ('" & Replace(strWord, "'", "''") & "')"
				m_objConn.Execute(sql)
				sql = "SELECT TOP 1 * FROM Words WHERE Word = '" & Replace(strWord, "'", "''") & "'"
				m_rs.Open sql, m_objConn, 0, 3
			End If

			AddWord = m_rs("ID")
			m_rs.Close

			m_objWordsDictionary.Add strWord, AddWord
		End If
	End Function

	Public Function AddURL(intTypeID, strURL, strTitle)
		' Add this URL/Title pair to the db if it's not there, return the URL ID
		' If strTitle = "" Then don't update title
		Dim sql, intURLID

		intURLID = GetURLID(strURL, intTypeID)
		If intURLID > -1 Then
			AddURL = intURLID
			Call Update(intURLID, intTypeID, strURL, strTitle)
		Else
			AddURL = Insert(intTypeID, strURL, strTitle)
		End If
	End Function

	Public Sub Update(intURLID, intTypeID, strURL, strTitle)
		Dim  sql

		sql = "UPDATE URLs SET "
		sql = sql & "Deleted = 0, "
		sql = sql & "TypeID = " & intTypeID & ", "
		sql = sql & "TotalTitleWords = 0, "
		sql = sql & "TotalCopyWords = 0 "
		If strTitle <> "" Then
			sql = sql & ", Title = '" & Replace(strTitle, "'", "''") & "' "
		End If
		If strURL <> "" Then
			sql = sql & ", URL = '" & Replace(strURL, "'", "''") & "' "
		End If

		sql = sql & " WHERE ID = " & intURLID
		m_objConn.Execute(sql)
	End Sub

	Public Function Insert(intTypeID, strURL, strTitle)
		Dim intURLID, bOverWrite, sql
		bOverWrite = GetNextURLID(intURLID)

		If bOverWrite Then
			'Use an existing "Deleted" record
			sql = "UPDATE URLs SET "
			sql = sql & "Deleted = 0, "
			sql = sql & "TypeID = " & intTypeID & ", "
			sql = sql & "TotalTitleWords = 0, "
			sql = sql & "TotalCopyWords = 0, "
			sql = sql & "URL = '" & Replace(strURL, "'", "''") & "', "
			If strTitle <> "" Then
				sql = sql & "Title = '" & Replace(strTitle, "'", "''") & "' "
			Else
				sql = sql & "Title = 'UNINDEXED' "
			End If

			sql = sql & " WHERE ID = " & intURLID
		Else
			'Add the URL
			sql = "INSERT INTO URLs(Deleted, TypeID, URL, Title) VALUES (" 
			sql = sql & "0, "
			sql = sql & intTypeID & ", "
			sql = sql & "'" & Replace(strURL, "'", "''") & "', "
			If strTitle <> "" Then
				sql = sql & "'" & Replace(strTitle, "'", "''") & "'"
			Else
				sql = sql & "'UNINDEXED'"
			End If
			sql = sql & ")"
		End If

		'Response.Write("sql=[" & sql & "]<BR>")
		m_objConn.Execute(sql)
		Insert = intURLID
	End Function

	Private Function GetURLID(strURL, intTypeID)
		Dim sql
		sql = "SELECT TOP 1 ID FROM URLs WHERE URL = '" & Replace(strURL, "'", "''") & "' AND TypeID=" & intTypeID
		m_rs.Open sql, m_objConn, 0, 3
		If Not m_rs.EOF Then
			GetURLID = m_rs("ID")
		Else
			GetURLID = -1
		End If
		m_rs.Close
	End Function

	Private Function GetNextURLID(intURLID)
		' RETURN True if this is an exisiting entry (i.e. we need to overwrite it)
		Dim sql

		sql = "SELECT TOP 1 ID FROM URLs WHERE Deleted = 1 ORDER BY ID"
		m_rs.Open sql, m_objConn, 0, 3
		If NOT m_rs.EOF Then
			GetNextURLID	= True
			intURLID		= m_rs("ID")
			m_rs.Close
			Exit Function
		Else
			m_rs.Close

			sql = "SELECT MAX(ID)+1 AS TheID FROM URLs"
			m_rs.Open sql, m_objConn, 0, 3
			If NOT m_rs.EOF Then
				GetNextURLID	= False
				intURLID		= m_rs("TheID")
				m_rs.Close
				Exit Function
			Else
				GetNextURLID	= False
				intURLID		= 1
				m_rs.Close
			End If
		End If
	End Function

	Public Sub ResetURL(intURLID)
		'Clear off all pre-existing words for this URL
		Dim sql

		sql = "DELETE FROM WordLink WHERE URLID = " & intURLID
		m_objConn.Execute(sql)
	End Sub

	Private Function GetTitle(strHTML)
		'Extract <TITLE> tag from HTML
		Dim intLeft, intRight, intPos
		
		GetTitle = ""
		intLeft = Instr(strHTML, "<TITLE>")
		If intLeft > 0 Then
			intLeft = intLeft + Len("<TITLE>")
			intRight = Instr(intLeft, strHTML, "</TITLE>")
			If intRight > 0 Then
				GetTitle = Mid(strHTML, intLeft, intRight-intLeft)
				
				intPos = InstrRev(GetTitle, ">")
				If intPos > 0 Then
					GetTitle = Right(GetTitle, Len(GetTitle)-(intPos+1))
					If Right(GetTitle, 1) = "." Then
						GetTitle = Left(GetTitle, Len(GetTitle)-1)
					End If
				End If
			End If
		End If
	End Function

	Public Function GetNextChunk(strChunkType, strText, intPos)
		Dim intEnd
		
		intPos = Instr(intPos, strText, "<!--" & strChunkType & "-->")
		If intPos > 0 Then
			intPos = intPos + Len("<!--" & strChunkType & "-->")
			intEnd = Instr(intPos, strText, "<!--/" & strChunkType & "-->")
			If intEnd = 0 Then
				intEnd = Len(strText)
			Else
				intEnd = intEnd-1 ' 1 back from the <
			End If

			GetNextChunk = Mid(strText, intPos, (intEnd - intPos)+1)

			intPos = intEnd + Len("<!--/" & strChunkType & "-->")
		Else
			intPos = 0
		End If
	End Function

	Public Function GetNextWord(strText, intPos)
		Dim intLen, strChar

		GetNextWord = ""
		intLen = Len(strText)
		
		' Find the start of the next word...
		Do While True
			If intPos > intLen Then Exit Do
			
			strChar = LCase(Mid(strText, intPos, 1))
			If strChar >= "a" And strChar <= "z" Then
				Exit Do
			End If

			intPos = intPos+1
		Loop

		If intPos > intLen Then 
			intPos = 0			' No more words
			Exit Function

		Else
			Do While True
				If intPos > intLen Then Exit Do
				
				strChar = LCase(Mid(strText, intPos, 1))
				If (strChar >= "a" And strChar <= "z") Or strChar = "'" Or strChar = "-" Then
					GetNextWord = GetNextWord & strChar
				Else
					Exit Do
				End If

				intPos = intPos+1
			Loop
			If Right(GetNextWord, 2) = "'s" then
				GetNextWord = Left(GetNextWord, Len(GetNextWord)-2)
			End If
			If Len(GetNextWord) > 30 Then
				GetNextWord = Left(GetNextWord, 30)
			End If
		End If
	End Function

	Public Function StripHTML(strBody)
		Dim strHTML, bFinished
		Dim aryDontFilter(2)
		Dim iStart, iEnd, intPos, intLeft, strReplaced
		
		StripHTML = strBody
		iStart 	= Instr(1, StripHTML, "<")
		iEnd	= Instr(1, StripHTML, ">")
		Do While iStart <> 0 And iEnd <> 0
			If iStart > iEnd Then 					'If '>' is found before '<' the string is corrupt.  '>' will be ignored
				iStart = iEnd + 1 					'adjust cursor to after '>' tag
			Else
				StripHTML = Left(StripHTML, iStart-1) & " " & Right(StripHTML, Len(StripHTML) - iEnd) 'remove tag
			End If
			intPos = iStart							'setting string cursor
			If intPos = 0 Then intPos = 1
			iStart 	= Instr(intPos, StripHTML, "<")
			iEnd	= Instr(intPos, StripHTML, ">")
		Loop

		If StripHTML <> "" Then
			' Next Remove HTML constants e.g. &NBSP;
			intPos = 1
			intLeft = 1
			strReplaced = ""
			bFinished = False
			Do While Not bFinished
				intPos = Instr(intPos, StripHTML, "&") - 1
				If intPos = -1 Then
					intPos = Len(StripHTML)
					bFinished = True
				End If
				strReplaced = strReplaced & Mid(StripHTML, intLeft, (intPos-intLeft)+1) & " "

				intLeft = Instr(intPos+1, StripHTML, ";")

				if intLeft = 0 Then
					intLeft = intPos + 2
				ElseIf (intLeft - intPos) > 6 Then
					intLeft = intPos + 2
				Else
					intLeft = intLeft + 1 ' Bypass the ;
				End If
				intPos = intLeft
			Loop

			StripHTML = strReplaced
		End If
	End Function

	Public Function RemoveDomain(strURL)
		' IN	= http://127.0.0.1/SiteSearch/history.asp
		' OUT	= SiteSearch/history.asp

		Dim intPos

		RemoveDomain = strURL

		intPos = Instr(strURL, "//")
		If intPos > 0 Then
			intPos = intPos + 2
			intPos = Instr(intPos, strURL, "/")
			if intPos > 0 Then
				RemoveDomain = Right(strURL, Len(strURL) - intPos)
			End If
		End If
	End Function

	Public Sub DeleteAll(intTypeID)
		' DELETEs ALL "URLs" of a particular type (use with care)
		Dim sql, intURLID

		sql = "SELECT ID FROM URLs WHERE TypeID=" & intTypeID
		m_rs.Open sql, m_objConn, 0, 3
		Do While Not m_rs.EOF
			intURLID = m_rs("ID")

			sql = "DELETE FROM WordLink WHERE URLID = " & intURLID
			m_objConn.Execute(sql)

			sql = "UPDATE URLs SET "
			sql = sql & "Deleted = 1, "
			sql = sql & "TypeID = 0, "
			sql = sql & "TotalTitleWords = 0, "
			sql = sql & "Title = 'DELETED', "
			sql = sql & "URL = '-', "
			sql = sql & "LastIndexed = GETDATE() "
			sql = sql & "WHERE ID = " & intURLID

			m_objConn.Execute(sql)

			m_rs.MoveNext
		Loop
		m_rs.Close
	End Sub
End Class
%>