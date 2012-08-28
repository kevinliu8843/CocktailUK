<%
Const AXIS_LEFT		= 1
Const AXIS_RIGHT	= 2

Class CBarChart
	Private m_aryData			' (Array of arrays)
	Private m_aryScaledData		' (Array of arrays)
	Private m_aryXLabels		' (Array of strings)
	Private m_bHasXLabels

	Private m_aryColors
	Private m_aryBarText		' Text for each bar element (Array of arrays)

	Private m_strFontNormal
	Private m_strFontSmall

	Private m_intNumBars

	Private m_intBarHeight
	Private m_intBarWidth
	Private m_intBarGapX

	Private m_strBarLineColor
	Private m_strBGColor

	Private m_intUpperLimit ' The top of the bar chart (in graph units)
	Private m_intChartHeight 'How high (in pixels) we want the chart to be scaled to
	Private m_intVerticalScale

	Private m_strTitle
	
	Private m_intYAxisLegendPositions
	Private m_strYAXISTitle

	Private m_strXAXISTitle
	Private m_strXAXISLabelMin
	Private m_strXAXISLabelMax

	Public Sub Class_Initialize()
		Call Reset()
	End Sub

	Public Sub Reset()
		ReDim m_aryData(0)
		ReDim m_aryBarText(0)
		ReDim m_aryXLabels(0)
		
		m_intNumBars = 0

		m_intBarWidth		= 16
		m_intBarGapX		= 5
		m_strBarLineColor	= "000000"
		m_strBGColor		= "CCCCCC"
		m_intChartHeight	= 200
		m_bHasXLabels		= False

		m_intYAxisLegendPositions = AXIS_LEFT
		m_strFontSmall		= "<font face=""Verdana, Arial"" size=""1"">"
		m_strFontNormal		= "<font face=""Verdana, Arial"" size=""2"">"
	End Sub

	Public Sub Class_Terminate()
	End Sub

	Public Sub SetTitle(strTitle)
		m_strTitle	= strTitle
	End Sub

	Public Sub SetBGColor(strColor)
		m_strBGColor = strColor
	End Sub

	Public Sub SetLineColor(strColor)
		m_strBarLineColor	= strColor
	End Sub

	Public Sub SetChartHeight(intHeight)
		m_intChartHeight = intHeight
	End Sub

	Public Sub SetXAxisTitle(strTitle)
		m_strXAXISTitle	= strTitle
	End Sub

	Public Sub SetXAxisMinAndMaxLabels(strMin, strMax)
		m_strXAXISLabelMin	= strMin
		m_strXAXISLabelMax	= strMax
	End Sub

	Public Sub SetYAxisTitle(strTitle)
		m_strYAXISTitle	= strTitle
	End Sub

	Public Sub SetYAxisLabelPos(intPos)
		m_intYAxisLegendPositions = intPos
	End Sub

	Public Sub SetBarWidth(intWidth)
		m_intBarWidth = intWidth
	End Sub

	Public Sub SetBarGap(intBarGapX)
		m_intBarGapX = intBarGapX
	End Sub

	Public Sub SetData(intColumn, aryData, strXLabel)
		If intColumn > UBound(m_aryData) Then
			ReDim Preserve m_aryData(intColumn)
			ReDim Preserve m_aryXLabels(intColumn)
			m_intNumBars = UBound(m_aryData)+1
		End If

		m_aryData(intColumn)	= aryData
		m_aryXLabels(intColumn) = strXLabel
		If strXLabel <> "" Then m_bHasXLabels = True
	End Sub

	Public Sub SetBarText(intColumn, aryText)
		If intColumn > UBound(m_aryBarText) Then
			ReDim Preserve m_aryBarText(intColumn)
		End If

		m_aryBarText(intColumn) = aryText
	End Sub

	Public Sub SetColors(aryColors)
		m_aryColors = aryColors
	End Sub

	Private Sub DetermineScale()
		Dim intMaxValue, i, j, intTotalBarHeight, aryMultiplier(2), arySteps(2)
		Dim intCurrent
		aryMultiplier(0) = 2
		aryMultiplier(1) = 2
		aryMultiplier(2) = 2.5

		arySteps(0)		= 5
		arySteps(1)		= 5
		arySteps(2)		= 4

		intMaxValue = 0
		For i = 0 to UBound(m_aryData)
			intTotalBarHeight = 0
			For j=0 To UBound(m_aryData(i))
				intTotalBarHeight = intTotalBarHeight + CLng(m_aryData(i)(j))
			Next
			If intTotalBarHeight > intMaxValue Then intMaxValue = intTotalBarHeight
		Next

		m_intUpperLimit	= 10
		intCurrent		= 1
		m_intBarHeight	= arySteps(intCurrent)
		Do
			If intMaxValue <= m_intUpperLimit Then Exit Do

			m_intUpperLimit = CLng(m_intUpperLimit * aryMultiplier(intCurrent))
			intCurrent = intCurrent + 1
			If intCurrent > 2 Then intCurrent = 0
			m_intBarHeight = arySteps(intCurrent)
		Loop

		'---------------------------------------------------------------
		m_intBarHeight = m_intUpperLimit / m_intBarHeight

		' Add an extra line at the top (for display purposes)
		'm_intUpperLimit = m_intUpperLimit + m_intBarHeight
		m_intVerticalScale = m_intChartHeight/m_intUpperLimit
	End Sub

	Private Sub ScaleData()
		Dim i,j, aryTemp
		ReDim m_aryScaledData(Ubound(m_aryData))
		For i = 0 to UBound(m_aryData)
			ReDim aryTemp(UBound(m_aryData(i)))
			m_aryScaledData(i) = aryTemp

			For j=0 To UBound(m_aryData(i))
				m_aryScaledData(i)(j) = CLng(m_aryData(i)(j) * m_intVerticalScale)
			Next
		Next
	End Sub

	Public Sub DrawGraph()
		Dim i
		Call DetermineScale()
		Call ScaleData()		' Generate Scaled values to stretch the graph

		Response.Write("<table cellspacing=""0"" cellpadding=""0"">")

		Call OutputTopOfChart()

		Response.Write("<tr>")

		If m_intYAxisLegendPositions And AXIS_LEFT Then
			Call OutputYAxisValues(AXIS_LEFT)
		Else		
			Call OutputBlankColumn(m_intBarGapX)
		End If
			
		For i=0 To m_intNumBars-1
			Call OutputColumn(i)
			Call OutputBlankColumn(m_intBarGapX)
		Next

		If m_intYAxisLegendPositions And AXIS_RIGHT Then
			Call OutputYAxisValues(AXIS_RIGHT)
		Else		
			Call OutputBlankColumn(m_intBarGapX)
		End If
	
		Response.Write("</tr>")

		Call OutputXAxisValues()

		Response.Write("</table>")
	End Sub

	Sub OutputTopOfChart()
		' Draw top of graph (first y axis marker + title )
		Response.Write("<tr>")
		Response.Write("<td align=""right"" bgcolor=""#" & m_strBGColor & """><BR>" & m_strFontSmall & m_strYAXISTitle & "<BR><BR>" & FormatNumber(m_intUpperLimit,0) & "&nbsp;</td>")
		Response.Write("<td align=""center"" colspan=""" & (m_intNumBars*2) & """ bgcolor=""#" & m_strBGColor & """>" & m_strFontSmall & m_strTitle & "</td>")
		Response.Write("<td bgcolor=""#" & m_strBGColor & """>&nbsp;</td>")
		Response.Write("</tr>")
		
		Response.Write("<tr><td colspan=""26"" bgcolor=""#" & m_strBarLineColor & """><img height=""1"" src=""spacer.gif""></td></tr>")
		' --------------------------------------------------------------------
	End Sub

	Private Sub OutputColumn(intColumn)
		Dim i, arySplit, intHeightDrawn, intIndex, intChunkHeight, strColor, strAltText

		arySplit = GetColumnChunks(intColumn)

		Response.Write("<td><table cellspacing=""0"" cellpadding=""0"">")
		
		intHeightDrawn = 0
		For i=UBound(arySplit) To 0 Step -2
			intIndex		= Int(arySplit(i-1))
			intChunkHeight	= CLng(arySplit(i))
			If intIndex = -1 Then
				' No bar
				strColor = m_strBGColor
			Else
				strColor = m_aryColors(intIndex)
			End If

			' Draw bar ---------------------------------------------
			If intChunkHeight > 0 Then
				strAltText = GetBarText(intColumn, intIndex)

				Response.Write("<tr>")
				If intIndex > -1 Then
					Response.Write("<td width=""1""><img width=""1"" height=""" & intChunkHeight & """ src=""black.gif""></td>")
					Response.Write("<td bgcolor=""#" & strColor & """><img width=""" & m_intBarWidth & """ height=""" & intChunkHeight & """ src=""spacer.gif"" alt=""" & strAltText & """></td>")
					Response.Write("<td width=""1""><img width=""1"" height=""" & intChunkHeight & """ src=""black.gif""></td>")
				Else
					Response.Write("<td bgcolor=""#" & m_strBGColor & """ width=""1""><img width=""1"" height=""" & intChunkHeight & """ src=""spacer.gif""></td>")
					Response.Write("<td bgcolor=""#" & strColor & """><img width=""" & m_intBarWidth & """ height=""" & intChunkHeight & """ src=""spacer.gif""></td>")
					Response.Write("<td bgcolor=""#" & m_strBGColor & """ width=""1""><img width=""1"" height=""" & intChunkHeight & """ src=""spacer.gif""></td>")
				End If
				Response.Write("</tr>")
			End If
			' ------------------------------------------------------

			intHeightDrawn = intHeightDrawn + intChunkHeight+1

			If intHeightDrawn >= CLng(m_intBarHeight*m_intVerticalScale) Then
				' Draw our bar height measure lines (50,100,150 etc.)
				Response.Write("<tr>")
				Response.Write("<td colspan=""3"" bgcolor=""#" & m_strBarLineColor & """><img width=""" & 1+m_intBarWidth+1 & """ height=""1"" src=""spacer.gif""></td>")
				Response.Write("</tr>")
				intHeightDrawn = 0
			Else
				' Draw a line above each bar section
				
				Response.Write("<tr>")
				Response.Write("<td><img width=""1"" height=""1"" src=""black.gif""></td>")
				Response.Write("<td><img width=""" & m_intBarWidth & """ height=""1"" src=""black.gif""></td>")
				Response.Write("<td><img width=""1"" height=""1"" src=""black.gif""></td>")
				Response.Write("</tr>")
			End If
		Next
		Response.Write("</table></td>")
	End Sub

	Private Function GetColumnChunks(intColumn)
		' For this column, calculate the chunks required to draw this graph and return it as an array.
		'  Copes with the scale bars that run through the graph too.
		
		Dim intIndex, intHeightDrawn, intBar, bFinished
		Dim intAmountOfCurrentBarDrawn, intCount, strParsed, intChunkHeight

		intIndex			= 0
		intHeightDrawn		= 0
		intBar				= CLng(m_intBarHeight*m_intVerticalScale)
		bFinished			= False
		intAmountOfCurrentBarDrawn = 0

		intCount = 0
		strParsed = ""
		Do While intCount < m_intUpperLimit 
			If bFinished=False Then 
				Do
					' Ignore zero values that wouldn't be displayed

					If CLng(m_aryScaledData(intColumn)(intIndex)) > 0 Then
						Exit Do
					End If

					IntIndex = IntIndex+1

					If intIndex > UBound(m_aryScaledData(intColumn)) Then 
						bFinished = True
						Exit Do
					End If
				Loop
			End If
			If bFinished = True Then
				intChunkHeight = intBar -intHeightDrawn
			Else
				If intHeightDrawn+(m_aryScaledData(intColumn)(intIndex)-intAmountOfCurrentBarDrawn) > intBar Then
					intChunkHeight = intBar -intHeightDrawn
				Else
					intChunkHeight = m_aryScaledData(intColumn)(intIndex)-intAmountOfCurrentBarDrawn
				End If
			End If

			intHeightDrawn = intHeightDrawn + intChunkHeight
			intAmountOfCurrentBarDrawn = intAmountOfCurrentBarDrawn + intChunkHeight
			
			If bFinished Then 
				If intChunkHeight > 0 Then
					strParsed = strParsed & -1 & "," & intChunkHeight-1 & ","
				End If
			Else
				If intChunkHeight > 0 Then
					strParsed = strParsed & intIndex & "," & intChunkHeight-1 & ","
				End If
			End If

			If bFinished = False Then
				If intAmountOfCurrentBarDrawn >= m_aryScaledData(intColumn)(intIndex) Then
					intAmountOfCurrentBarDrawn = 0
					intIndex=intIndex+1
					If intIndex > UBound(m_aryScaledData(intColumn)) Then bFinished = True
				End If
			End If

			If intHeightDrawn = intBar Then
				intCount = intCount + m_intBarHeight
				intBar = intBar + CLng(m_intBarHeight*m_intVerticalScale)
			End If
		Loop
		
		strParsed = Left(strParsed, Len(strParsed)-1)

		GetColumnChunks = Split(strParsed, ",")
	End Function

	Private Sub OutputBlankColumn(intWidth)
		Dim intCount, intHeight
		intCount	= m_intUpperLimit
		intHeight	= CLng(m_intBarHeight*m_intVerticalScale)-1

		Response.Write("<td><table cellspacing=""0"" cellpadding=""0"">")
		Do While intCount > 0
			Response.Write("<tr><td bgcolor=""#" & m_strBGColor & """ height=""" & intHeight & """ width=""" & intWidth & """><img src=""spacer.gif""></td></tr>")
			Response.Write("<tr><td bgcolor=""#" & m_strBarLineColor & """><img width=""" & intWidth & """ height=""1"" src=""spacer.gif""></td></tr>")

			intCount = intCount - m_intBarHeight

			Response.Write("</td>")
		Loop
		Response.Write("</table></td>")
	End Sub

	Private Sub OutputYAxisValues(intPosition)
		Dim intValue, intHeight, strOutput
		intValue	= m_intUpperLimit
		intHeight	= CLng(m_intBarHeight*m_intVerticalScale)-1

		Response.Write("<td bgcolor=""#" & m_strBGColor & """><table width=""100%"" cellspacing=""0"" cellpadding=""0"">")
		Do While intValue >0
			Response.Write("<tr><td align=""right"" valign=""bottom"" height=""" & intHeight & """ bgcolor=""#" & m_strBGColor & """>")
			strOutput = FormatNumber(intValue-m_intBarHeight,0)
			If intPosition = AXIS_LEFT Then
				strOutput = "&nbsp;" & strOutput & "&nbsp;"
			End If

			Response.Write(m_strFontSmall & strOutput & "</font></td></tr>")

			Response.Write("<tr><td bgcolor=""#" & m_strBarLineColor & """><img height=""1"" src=""spacer.gif""></td></tr>")

			intValue = intValue - m_intBarHeight
			Response.Write("</td>")
		Loop
		Response.Write("</table></td>")
	End Sub

	Private Function GetBarText(intColumn, intIndex)
		Dim bFound

		If intIndex > -1 Then
			bFound = False
			If intColumn <= UBound(m_aryBarText) Then
				If IsArray(m_aryBarText(intColumn)) Then
					If intIndex<=UBound(m_aryBarText(intColumn)) Then
						GetBarText = m_aryBarText(intColumn)(intIndex)
						bFound = True
					End If
				End If
			End If
			If bFound = False Then
				GetBarText = m_aryData(intColumn)(intIndex)
			End If
		End If
	End Function

	Sub OutputXAxisValues()
		Dim intLeft, intRight, i

		If m_bHasXLabels Then
			Response.Write("<tr bgcolor=""#" & m_strBGColor & """><td><img src=""spacer.gif""></td>")

			For i=0 To m_intNumBars-1
				Response.Write("<td align=""center"">" & m_strFontSmall & m_aryXLabels(i) & "</font></td>")
				Response.Write("<td><img src=""spacer.gif""></td>")
			Next

			Response.Write("<td><img src=""spacer.gif""></td></tr>")
		End If

		If m_strXAXISLabelMin <> "" Or m_strXAXISLabelMax <> "" Then
			intLeft		= (m_intNumBars*2)\2
			intRight	= ((m_intNumBars*2) - intLeft)-1

			Response.Write("<tr bgcolor=""#" & m_strBGColor & """><td><img src=""spacer.gif""></td>")
			Response.Write("<td align=""left"" colspan=""" & intLeft & """>" & m_strFontSmall & m_strXAXISLabelMin & "</font></td>")
			Response.Write("<td align=""right"" colspan=""" & intRight & """>" & m_strFontSmall & m_strXAXISLabelMax & "</font></td>")
			Response.Write("<td colspan=""2""><img src=""spacer.gif""></td></tr>")
		End If

		
		Response.Write("<tr bgcolor=""#" & m_strBGColor & """><td><img src=""spacer.gif""></td>")
		Response.Write("<td align=""center"" colspan=""" & m_intNumBars*2 & """>" & m_strFontSmall & m_strXAXISTitle & "</font></td>")
		Response.Write("<td><img src=""spacer.gif""></td></tr>")
	End Sub
End Class

Sub OutputChart(strType, strTitle)
	Dim i, intMonth, intYear, intBarValue, intCocktailViews
	Dim intMinYear, intMaxYear, strBarText1, strBarText2, intProjectedTarget, intProjectedTargetDiff
	Dim dateTemp

	' Get Data for chart ---------------------------------------------------------------
	strSQL =			"SELECT * FROM Counter "
	'strSQL = strSQL &	"WHERE DateSerial(Year, Month, 1) > DateSerial(" & Year(dateDate)-1 & ", " & Month(dateDate) & ", 1) "
	'strSQL = strSQL &	" AND DateSerial(Year, Month, 1) <= DateSerial(" & Year(dateDate) & ", " & Month(dateDate) & ", 1) "
	strSQL = strSQL &	"ORDER BY Year, Month"

	' Populate all bars with basic information (for bars that don't exist in DB) ------------
	dateTemp = CDate("1/" & MonthName(Month(dateDate), True) & "/" & Year(dateDate))
	For i=11 To 0 Step-1
		If i=0 Then
			intMinYear = Year(dateTemp)
		ElseIf i=11 Then
			intMaxYear = Year(dateTemp)
		End If

		Call objBarChart.SetData(i, Array(0), Left(MonthName(Month(dateTemp), True), 1))
		dateTemp = DateAdd("m", -1, dateTemp)
	Next
	'----------------------------------------------------------------------------------------

	rs.Open strSQL, cn, 0, 3
	iStep = 0
	Do While Not rs.EOF
		If (DateSerial(Int(rs("Year")), Int(rs("Month")), 1) > DateSerial(Year(dateDate)-1, Month(dateDate), 1) AND DateSerial(Int(rs("Year")), Int(rs("Month")), 1) <= DateSerial(Year(dateDate), Month(dateDate), 1)) Then
			iStep = iStep + 1
			intMonth	= Int(rs("Month"))
			intYear		= Int(rs("Year"))
			intBarValue	= CLng(rs(strType))
	
			' Calculate index of bar based on the date received...
			dateTemp = CDate("1/" & MonthName(intMonth, True) & "/" & intYear)
			i = DateDiff("m", dateDate, dateTemp)+11
			' --------------------------------------------------------
	
			strBarText1 = FormatNumber(intBarValue,0) & " " & strType & VBCrLf & MonthName(intMonth, False) & " " & intYear
			
			If intMonth = Int(Month(dateNow)) And intYear = Int(Year(dateNow)) Then
				intProjectedTarget = CalculateProjectedTarget(intBarValue)
				intProjectedTargetDiff = intProjectedTarget - intBarValue
	
				strBarText2 = "Projected " & FormatNumber(intProjectedTarget,0) & " " & strType & VBCrLf & MonthName(intMonth, False) & " " & intYear
	
				Call objBarChart.SetData(i, Array(intBarValue, intProjectedTargetDiff), Left(MonthName(intMonth, True), 1))
				Call objBarChart.SetBarText(i, Array(strBarText1, strBarText2))
	
			Else
				Call objBarChart.SetData(i, Array(intBarValue), Left(MonthName(intMonth, True), 1))
				Call objBarChart.SetBarText(i, Array(strBarText1))
			End If
		End If
		rs.MoveNext
	Loop
	rs.Close

	'------------------Projection-------------------------------
	if (Year(Now()) <= Year(CDate(request("dtDate")))) Then
	strSQL =			"SELECT * FROM Counter "
	'strSQL = strSQL &	"WHERE DateSerial(Year, Month, 1) > DateSerial(" & Year(dateNow)-1 & ", " & Month(dateNow) & ", 1) "
	'strSQL = strSQL &	" AND DateSerial(Year, Month, 1) <= DateSerial(" & Year(dateNow) & ", " & Month(dateNow) & ", 1) "
	strSQL = strSQL &	"ORDER BY Year, Month"

	rs.open strSQL , cn, 0, 3
	For i=1 to 12
		'If (DateSerial(Int(rs("Year")), Int(rs("Month")), 1) > DateSerial(Year(dateDate)-1, Month(dateDate), 1) AND DateSerial(Int(rs("Year")), Int(rs("Month")), 1) <= DateSerial(Year(dateDate), Month(dateDate), 1)) Then
			if (i=1 or i=2 or i=3) then
				dblAverage1 = dblAverage1 + CLng(rs(strType))
			end if
			if(i=10 or i=11 or i=12) then
				dblAverage2 = dblAverage2 + CLng(rs(strType))
			end if
			if (i=12) then
				intProjectedTarget = CalculateProjectedTarget(CLng(rs(strType)))
			end if
		'End If
		rs.MoveNext
	Next
	rs.Close
	dblAverage1 = dblAverage1 / 3
	dblAverage2 = dblAverage2 / 3
	dblDistance = 12 - 4
	dblGradient = ((dblAverage2-dblAverage1) / dblDistance)

	dateTemp = CDate(request("dtDate"))
	For i=iStep To 11
		If i=0 Then
			intMinYear = Year(dateTemp)
		ElseIf i=11 Then
			intMaxYear = Year(dateTemp)
		End If

		intProj = dblGradient*(DateDiff("m", Now(), dateTemp)-5) + intProjectedTarget
		
		Call objBarChart.SetData(i, Array(0, intProj), Left(MonthName(Month(DateAdd("m", -5, dateTemp)), True), 1))
		strBarText1 = FormatNumber(intProj ,0) & " " & strType & VBCrLf & MonthName(Month(DateAdd("m", -5, dateTemp)), False) & " " & intYear
		strBarText2 = "Projected " & FormatNumber(intProj ,0) & " " & strType & VBCrLf & MonthName(Month(DateAdd("m", -5, dateTemp)), False) & " " & Year(DateAdd("m", -5, dateTemp))
		Call objBarChart.SetBarText(i, Array(strBarText1, strBarText2))
		dateTemp = DateAdd("m", 1, dateTemp)
	Next
	'----------------------------------------------------------------------------------------
	end if
	
	' Set up chart properties ------------------------------------------------------------

	Call objBarChart.SetTitle(strTitle)
	Call objBarChart.SetXAxisMinAndMaxLabels(intMinYear, intMaxYear)
	Call objBarChart.SetXAxisTitle("Month")
	Call objBarChart.SetYAxisTitle(strType)
	Call objBarChart.SetYAxisLabelPos(AXIS_LEFT)
	Call objBarChart.SetBGColor("F1F1F1")
	Call objBarChart.SetLineColor("000000")
	Call objBarChart.SetColors(Array("636388", "CCCCDD"))
	'------------------------------------------------------------------------------------

	Call objBarChart.DrawGraph()
End Sub

%>