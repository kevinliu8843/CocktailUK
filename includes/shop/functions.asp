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
Sub GetEstimatedDelivery(cn, rs, dte, blnIsDispatchDate, priority, saturdaydeliveries, delay1, delay2, dtePack, dteDispatch, dte1, dte2)
	Dim aryNPDays, dteStart, strDispatchTime, intBypassed
	
	dte1 = dte 
	dte2 = dte
	
	aryNPDays = GetNonWorkingDays(cn, rs)
	
	If blnIsDispatchDate Then
		dteStart = CDate(MediumDate(dte) & " 09:00:00")
	Else
		dteStart = dte
	End If
	
	Call GetPackAndDispatchDatesFromDate(cn, rs, dteStart, priority, intBypassed, dtePack, dteDispatch)
	
	'Calculate the date ranges
	dte1 = GetWorkingDay(dteDispatch, delay1, aryNPDays, intBypassed, saturdaydeliveries)
	dte2 = GetWorkingDay(dteDispatch, delay2, aryNPDays, intBypassed, saturdaydeliveries)
End Sub

Sub GetPackAndDispatchDatesFromDate(cn, rs, dteStart, priority, intBypassed, dtePack, dteDispatch)
	Dim aryNPDays, strDispatchTime, strStartWorkTime, dblPackDelayHrs, strUWarehousecutOffTime
	
	aryNPDays = GetNonWorkingDays(cn, rs)
	
	strStartWorkTime 		= "09:00:00"
	strDispatchTime  		= "17:00:00"
	strUWarehousecutOffTime		= "16:30:00"
	If WeekDay(Now()) = vbFriday Then
		strUWarehousecutOffTime	= "15:30:00"
	End If

	dteStart = GetWorkingDay(dteStart, 0, aryNPDays, intBypassed, False)
	If intBypassed > 0 then
		dteStart = CDate(MediumDate(dteStart) & " " & strStartWorkTime)
	End If
	
	If priority Then 
		dblPackDelayHrs = 0
	Else
		dblPackDelayHrs = 0
	End If
	
	'Add on a delay between start date and packing
	dtePack = DateAdd("h", dblPackDelayHrs, dteStart) 
	
	'If the predicted pack time is after the cut off time for packing, move it to the next working day
	If dtePack > CDate(MediumDate(dtePack) & " " & strUWarehousecutOffTime) Then
		dtePack = GetWorkingDay(dtePack, 1, aryNPDays, intBypassed, False)
		dtePack = CDate(MediumDate(dtePack) & " " & strStartWorkTime)
	End If
	
	'Estimate dispatch date
	dteDispatch = CDate(MediumDate(dtePack) & " " & strUWarehousecutOffTime) ' Make this date the last order date/time instead of dispatch time
'response.write dtePack & " " & dteDispatch & "<BR>"
End Sub

Function GetNextWorkingDay(cn, rs)
	Dim intBypassed
	GetNextWorkingDay = GetWorkingDay(Now(), 0, GetNonWorkingDays(cn, rs), intBypassed, False)
End Function

Function GetNextWorkingDaySPDate(cn, rs, dte)
	Dim intBypassed
	GetNextWorkingDay = GetWorkingDay(dte, 0, GetNonWorkingDays(cn, rs), intBypassed, False)
End Function

Function GetPreviousWorkingDay(cn, rs)
	Dim intBypassed
	GetPreviousWorkingDay = GetWorkingDay(Now(), -1, GetNonWorkingDays(cn, rs), intBypassed, False)
End Function

Function GetPreviousWorkingDaySPDate(cn, rs, dte)
	Dim intBypassed
	GetPreviousWorkingDaySPDate = GetWorkingDay(dte, -1, GetNonWorkingDays(cn, rs), intBypassed, False)
End Function

Function GetWorkingDayOrder(cn, rs, dte, intDays)
	Dim intBypassed
	GetWorkingDayOrder = GetWorkingDay(dte, intDays, GetNonWorkingDays(cn, rs), intBypassed, false)
End Function

Function GetNonWorkingDays(cn, rs)
	Dim aryNPDays
	
	rs.open "SELECT dte FROM jointnonpackingdays WHERE DATEDIFF(day, dte, GETDATE()) <= 30 ORDER BY dte", cn
	If NOT rs.EOF Then
		aryNPDays = rs.GetRows()
	Else
		ReDim aryNPDays(-1, -1)
	End If
	rs.close
	
	GetNonWorkingDays = aryNPDays
End Function

Function GetWorkingDay(dte, intDays, aryNPDays, intBypassed, saturdaysallowed)
	Dim dteNewDate, intDaysFound, intLimit, intProgression
 
	intLimit 		= 100
	dteNewDate 		= dte
	intDaysFound	= -1 
	intBypassed 	= 0

	While intDaysFound < Abs(intDays) AND intBypassed < intLimit
		If IsWorkingDay(dteNewDate, aryNPDays, saturdaysallowed) Then
			intDaysFound = intDaysFound + 1
		Else
			intBypassed = intBypassed + 1
		End If
		If intDaysFound < Abs(intDays) Then
			If intDays >= 0 Then
				intProgression = 1
			Else
				intProgression = -1
			End If
			dteNewDate =  DateAdd("d", intProgression, dteNewDate)
		End If
	Wend

	GetWorkingDay = dteNewDate
End Function

Function IsWorkingDay(dte, aryNPDays, saturdaysallowed)
	Dim i, strDate

	If (Weekday(dte) = vbSaturday AND NOT saturdaysallowed) OR Weekday(dte) = vbSunday then
		IsWorkingDay = False 
		Exit Function
	End If

	If IsArray(aryNPDays) Then
		strDate = MediumDate(dte)
		For i=0 To UBound(aryNPDays, 2)
			If strDate = MediumDate(aryNPDays(0, i)) Then
				IsWorkingDay = False
				Exit Function
			End If
		Next
	End If
	
	IsWorkingDay = True
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