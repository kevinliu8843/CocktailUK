<%
Dim strCriticalErrorMessage
Dim blnCriticalErrors
strCriticalErrorMessage = ""
blnCriticalErrors = False

Sub TrapErrors()
	TrapError Err.description, True
End Sub

Sub CovertTrapErrors()
	TrapError Err.description, False
End Sub

Sub TrapError(strError, blnMessage)
	blnCriticalErrors = True
	strCriticalErrorMessage = strCriticalErrorMessage & strError
	
	Dim objError
	Set objError = Server.GetLastError()
	If NOT objError Is Nothing Then
		If Len(CStr(objError.ASPCode)) > 0 Then 
			strCriticalErrorMessage = strCriticalErrorMessage & "IIS Error Number: " & objError.ASPCode & "<BR>"
		End If
		If objError.Number > 0 Then 
			strCriticalErrorMessage = strCriticalErrorMessage & "COM Error Number: " & objError.Number & "<BR>"
		End If
		If Len(CStr(objError.Source )) > 0 Then 
			strCriticalErrorMessage = strCriticalErrorMessage & "Error Source: " & objError.Source & "<BR>"
		End If
		If Len(CStr(objError.File )) > 0 Then 
			strCriticalErrorMessage = strCriticalErrorMessage & "File Name: " & objError.File & "<BR>"
		End If
		If objError.Line > 0 Then 
			strCriticalErrorMessage = strCriticalErrorMessage & "Line Number: " & objError.Line & "<BR>"
		End If
		If Len(CStr(objError.Description )) > 0 Then 
			strCriticalErrorMessage = strCriticalErrorMessage & "Brief Description: " & objError.Description & "<BR>"
		End If
		If Len(CStr(objError.ASPDescription )) > 0 Then 
			strCriticalErrorMessage = strCriticalErrorMessage & "Full Description: " & objError.ASPDescription & "<BR>"
		End If
		Set objError = Nothing
		
		If strCriticalErrorMessage <> "" Then
			Call ProcessErrors(blnMessage)
		End If
	End If
End Sub

'If there are any errors, this function will email tech. support
Sub ProcessErrors(blnMessage)
	Dim strErrorEmailBody, Item, QS, RF
	If blnCriticalErrors Then
		strErrorEmailBody = "<LINK rel=""stylesheet"" type=""text/css"" href=""/style/admin/admin.css"">"
		strErrorEmailBody = strErrorEmailBody & "<P>At " & Now & " the following error occurred:<br>" & strCriticalErrorMessage & "</p>"

		'*** Display Dynamic Session Variables ***
		strErrorEmailBody = strErrorEmailBody & "<table cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse;"" class=""partadmintable"" style=""margin-top: 10px"">"& VbCrLf
		strErrorEmailBody = strErrorEmailBody & "    <tr>"& VbCrLf
		strErrorEmailBody = strErrorEmailBody & "      <th colspan=""2"" style=""text-align: left"" width=""200"">"& VbCrLf
		strErrorEmailBody = strErrorEmailBody & "		<b>Session Variables</b></th>"& VbCrLf
		strErrorEmailBody = strErrorEmailBody & "    </tr>"& VbCrLf
		strErrorEmailBody = strErrorEmailBody & "    <tr>"& VbCrLf
		strErrorEmailBody = strErrorEmailBody & "      <th style=""text-align: left"">Variable Name</th>"& VbCrLf
		strErrorEmailBody = strErrorEmailBody & "      <th style=""text-align: left"">Data</th>"& VbCrLf
		strErrorEmailBody = strErrorEmailBody & "    </tr>"& VbCrLf
		For Each Item In Session.contents
			If IsArray(Session(Item)) Then
				strErrorEmailBody = strErrorEmailBody & "<tr><td valign=""top"">" & Item & "</td><td valign=""top"">" & PrintArray(Session(Item)) & "</td></tr>" & VbCrLf
			Else
				strErrorEmailBody = strErrorEmailBody & "<tr><td valign=""top"">" & Item & "</td><td valign=""top"">" & Session(Item) & "</td></tr>" & VbCrLf
			End If
		Next
		strErrorEmailBody = strErrorEmailBody & "</table>"
		
		iF Request.QueryString.Count > 0 Then
			'*** Display QueryString Variables ***
			strErrorEmailBody = strErrorEmailBody & "<table cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse;"" class=""partadmintable"" style=""margin-top: 10px"">"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "    <tr>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "      <th colspan=""2"" style=""text-align: left"" width=""200"">"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "		<b>QueryString Collection</b></th>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "    </tr>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "    <tr>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "      <th style=""text-align: left"">Variable Name</th>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "      <th style=""text-align: left"">Data</th>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "    </tr>"& VbCrLf
			For Each Item In Request.QueryString
				strErrorEmailBody = strErrorEmailBody & "<tr><td valign=""top"">" & Item & "</td><td valign=""top"">" & Request.QueryString(Item) & "</td></tr>" & VbCrLf
			Next
			strErrorEmailBody = strErrorEmailBody & "</table>"
		End If

		sContentType = Request.ServerVariables("HTTP_CONTENT_TYPE")	
		if InStr(sContentType,"multipart/form-data")=0 Then
			iF Request.Form.Count > 0 Then
				'*** Display POST Form Variables ***
				strErrorEmailBody = strErrorEmailBody & "<table cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse;"" class=""partadmintable"" style=""margin-top: 10px"">"& VbCrLf
				strErrorEmailBody = strErrorEmailBody & "    <tr>"& VbCrLf
				strErrorEmailBody = strErrorEmailBody & "      <th colspan=""2"" style=""text-align: left"" width=""200"">"& VbCrLf
				strErrorEmailBody = strErrorEmailBody & "		<b>Form Collection</b></th>"& VbCrLf
				strErrorEmailBody = strErrorEmailBody & "    </tr>"& VbCrLf
				strErrorEmailBody = strErrorEmailBody & "    <tr>"& VbCrLf
			    	strErrorEmailBody = strErrorEmailBody & "      <th style=""text-align: left"">Variable Name</th>"& VbCrLf
			    	strErrorEmailBody = strErrorEmailBody & "      <th style=""text-align: left"">Data</th>"& VbCrLf
			    	strErrorEmailBody = strErrorEmailBody & "    </tr>"& VbCrLf
				For Each Item In Request.Form
					strErrorEmailBody = strErrorEmailBody & "<tr><td valign=""top"">" & Item & "</td><td valign=""top"">" & Request.Form(Item) & "</td></tr>" & VbCrLf
				Next
				strErrorEmailBody = strErrorEmailBody & "</table>"
			End If
		End If
		
		If Request.ServerVariables.Count > 0 Then
			'*** Display ServerVariables Variables ***
			strErrorEmailBody = strErrorEmailBody & "<table cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse;"" class=""partadmintable"" style=""margin-top: 10px"">"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "    <tr>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "      <th colspan=""2"" style=""text-align: left"" width=""200"">"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "		<b>ServerVariables Collection</b></th>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "    </tr>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "    <tr>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "      <th style=""text-align: left"">Variable Name</th>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "      <th style=""text-align: left"">Data</th>"& VbCrLf
			strErrorEmailBody = strErrorEmailBody & "    </tr>"& VbCrLf
			For Each Item In Request.ServerVariables
				strErrorEmailBody = strErrorEmailBody & "<tr><td valign=""top"">" & Item & "</td><td valign=""top"">" & Request.ServerVariables(Item) & "</td></tr>" & VbCrLf
			Next
			strErrorEmailBody = strErrorEmailBody & "</table>"
		End If
		
		If NOT Session("admin") Then
			Call SendEmail("theteam@cocktail.uk.com", "theteam@cocktail.uk.com", "", "", "Cocktail : UK Error", strErrorEmailBody, True, "")
		Else
			Response.write "<div style=""border: 2px dashed red; padding: 5px;""><h3 style=""color: red"">Error Report</h3>" & strErrorEmailBody & "</DIV>"
		End If

		If blnMessage AND NOT Session("admin") Then
			Response.Write "<p align=""center""><font color=red><b>There has been an internal script error on this page. Technical Support " & _
		                   "has already been notified. Thank you for your patience.</b></font></p>"
		End If
	End If
End Sub

Function PrintArray(aryArray)
	On Error Resume Next
	Dim i, j, k, strOut, strElement, aryDimensions(10), strDimensions
	i=0
	strDimensions = ""
	For Each strElement in aryArray
		i = i + 1
	Next
	j = 0
	k=1
	Do While j >= 0 AND k < 10
		j = UBound(aryArray, k)
		If j > 0 Then
			strDimensions = strDimensions & j & ", "
			aryDimensions(k-1) = j
		End If
		j = 0
		k = k + 1
	Loop
	If strDimensions <> "" Then
		strDimensions = Left(strDimensions, Len(strDimensions)-2)
	End If
	strOut = "Array ("&strDimensions&"):<BR>"
	j=0
	For Each strElement in aryArray
		strOut = strOut & strElement
		j = j + 1
		If j = aryDimensions(0)+1 Then
			strOut = strOut & "<BR>"
			j=0
		Else
			strOut = strOut & ",&nbsp;"
		End If
	Next	
	PrintArray = strOut
End Function
%>
