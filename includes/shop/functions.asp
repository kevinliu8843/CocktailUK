<%
Function stripHTML(strHTML)
'Strips the HTML tags from strHTML

  Dim objRegExp, strOutput
  Set objRegExp = New Regexp

  objRegExp.IgnoreCase = True
  objRegExp.Global = True
  objRegExp.Pattern = "<(.|\n)+?>"

  'Replace all HTML tag matches with the empty string
  strOutput = objRegExp.Replace(strHTML, "")
  
  stripHTML = strOutput    'Return the value of strOutput

  Set objRegExp = Nothing
End Function

Function PrettyDateShort(dte)
	Dim strOut, strSuperScript

	If IsNull(dte) Then
		Exit Function
	End If

	If CDate(dte) < CDate("1-Jan-1950") Then
		Exit Function
	End If
	
	SELECT Case Int(Day(dte))
		Case 1,21,31	strSuperScript = "st"
		Case 2,22     	strSuperScript = "nd"
		Case 3,23     	strSuperScript = "rd"
		Case Else    	strSuperScript = "th"
	End Select
	
	strOut = Day(dte) & strSuperScript & "&nbsp;" & MonthName(Month(dte), True)
	If Year(Now()) <> Year(dte) Then
		strOut = strOut & "&nbsp;'" & Right(Year(dte), 2)
	End If
	
	PrettyDateShort = strOut
End Function

Function PrettyDate(dte)
	Dim strOut, strSuperScript

	If dte = "" Or dte = "NULL" Or IsNull(dte) Then 
		Exit Function
	End If

	If CDate(dte) < CDate("1-Jan-1950") Then
		Exit Function
	End If
	
	SELECT Case Int(Day(dte))
		Case 1,21,31	strSuperScript = "st"
		Case 2,22     	strSuperScript = "nd"
		Case 3,23     	strSuperScript = "rd"
		Case Else    	strSuperScript = "th"
	End Select
	strOut = Weekdayname(Weekday(dte), True) & "&nbsp;" & Day(dte) & strSuperScript & "&nbsp;" & MonthName(Month(dte))
	If Year(Now()) <> Year(dte) Then
		strOut = strOut & "&nbsp;'" & Right(Year(dte), 2)
	End If
	PrettyDate = strOut
End Function

Function PrettyDateTime(dte)
	Dim strOut

	If dte = "" Or dte = "NULL" Or IsNull(dte) Then 
		Exit Function
	End If

	If CDate(dte) < CDate("1-Jan-1950") Then
		Exit Function
	End If
	
	strOut = PrettyDate(dte)
	strOut = strOut & " " & FormatDateTime(dte, 3)
	If Right(strOut, 3) = ":00" Then
		strOut = Left(strOut, Len(strOut)-3)
	End If
	PrettyDateTime = strOut
End Function

Function PrettyDateShortTime(dte)
	Dim strOut

	If dte = "" Or dte = "NULL" Or IsNull(dte) Then 
		Exit Function
	End If

	If CDate(dte) < CDate("1-Jan-1950") Then
		Exit Function
	End If
	
	strOut = PrettyDateShort(dte)
	strOut = strOut & " " & FormatDateTime(dte, 3)
	If Right(strOut, 3) = ":00" Then
		strOut = Left(strOut, Len(strOut)-3)
	End If
	PrettyDateShortTime = strOut
End Function
%>