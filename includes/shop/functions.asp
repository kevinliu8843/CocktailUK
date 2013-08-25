<!--#include virtual="/includes/shop/currency.asp" -->
<%
Function IsCollectionOnly(cn, rs, id)
	Dim blnCollect
	blnCollect = False
	'Is it collection only?
	strSQL = "SELECT delID from dsprodallowdelivery WHERE prodID=" & id
	rs.open strSQL, cn, 0, 3
	If NOT rs.EOF Then
		If rs("delID")=  -1 Then
			blnCollect = True
		End If
	End If
	rs.close
	IsCollectionOnly = blnCollect
End Function

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

' Examples for clarity
' =============================
' Net + VAT = Gross
' =============================
' VAT Rate:       17.5%
' Price Incl VAT: £9.99 (Gross)
' VAT:            £1.49 (VAT)
' Net:            £8.50 (Net)
' =============================
' VAT Rate:       15%
' Price Incl VAT: £9.99 (Gross)
' VAT:            £1.30 (VAT)
' Net:            £8.69 (Net)
' =============================

Function CalculateVATFromGross(dblGross, dblVatRate)
	CalculateVATFromGross = Round(dblGross - (dblGross/ (1 + (dblVatRate/100))), 2)
End Function

Function CalculateNetFromGross(dblGross, dblVatRate)
	CalculateNetFromGross = Round(dblGross/ (1 + (dblVatRate/100)), 2)
End Function

Function CalculateVATFromNet(dblNet, dblVatRate)
	alculateVATFromNet = Round(dblNet * dblVatRate/100, 2)
End Function

Function CalculateGrossFromNet(dblNet, dblVatRate)
	CalculateGrossFromNet = Round(dblNet * (1 + (dblVatRate/100)), 2)
End Function
%>