<%
Private Sub GetDrinkstuffSales(dArrTrueSales, iArrTrueSales, dArrTrueProfits)
	call GetSalesActual(dArrTrueSales, iArrTrueSales, dArrTrueProfits, "SALES", "http://admin.drinkstuff.com/affiliate/dsxml.asp")
End Sub

Private Sub GetBarmansSales(dArrSales, iArrVolume, dArrTrueProfits)
	call GetSalesActual(dArrSales, iArrVolume, dArrTrueProfits, "SALES", "http://admin.barmans.co.uk/affiliate/barxml.asp")
End Sub

Private Sub GetAffiliateSales(dArrSales, iArrVolume, dArrTrueProfits)
	call GetSalesActual(dArrSales, iArrVolume, dArrTrueProfits, "SALES", "http://admin.drinkstuff.com/affiliate/login.asp?user=10724&pass=leetracey&xml=true")
End Sub

Private Sub GetCEAffiliateSales(dArrSales, iArrVolume, dArrTrueProfits)
	call GetSalesActual(dArrSales, iArrVolume, dArrTrueProfits, "SALES", "http://admin.drinkstuff.com/affiliate/login.asp?user=20724&pass=leetracey&xml=true")
End Sub

Private Sub GetBarmansAffiliateSales(dArrSales, iArrVolume, dArrTrueProfits)
	call GetSalesActual(dArrSales, iArrVolume, dArrTrueProfits, "SALES", "http://admin.barmans.co.uk/affiliate/login.asp?user=10724&pass=leetracey&xml=true")
End Sub

Public Sub GetSalesActual(dArrTrueSales, iArrTrueSales, dArrTrueProfits, strRoot, strXmlDoc)
	On Error Resume Next
	Dim objXmlDoc, objXmlHttpCat, strXML
	Dim strSales1, intSales1, strSales2, intSales2, strSales3, intSales3, strSales4, intSales4
	Dim dblSales1, dblSales2, dblSales3, dblSales4

	Set objXmlHttpCat= Server.CreateObject("MSXML2.ServerXMLHTTP") 
	objXmlHttpCat.open "GET", strXmlDoc, False
	objXmlHttpCat.setTimeouts 5000, 5000, 5000, 5000
	objXmlHttpCat.send ""
	Set objXmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	strXML = objXmlHttpCat.responseText
	Set objXmlHttpCat = nothing
	objXmlDoc.loadXML(strXML)
	If strXML = "" Then
		Exit Sub
	End If
	
	Set strSales1 = objXmlDoc.getElementsByTagName(strRoot&"/TODAY/SALES")
	Set intSales1 = objXmlDoc.getElementsByTagName(strRoot&"/TODAY/VOLUME")
	Set dblSales1 = objXmlDoc.getElementsByTagName(strRoot&"/TODAY/COMMISSION")
	Set strSales2 = objXmlDoc.getElementsByTagName(strRoot&"/LASTMONTH/SALES")
	Set intSales2 = objXmlDoc.getElementsByTagName(strRoot&"/LASTMONTH/VOLUME")
	Set dblSales2 = objXmlDoc.getElementsByTagName(strRoot&"/LASTMONTH/COMMISSION")
	Set strSales3 = objXmlDoc.getElementsByTagName(strRoot&"/YESTERDAY/SALES")
	Set intSales3 = objXmlDoc.getElementsByTagName(strRoot&"/YESTERDAY/VOLUME")
	Set dblSales3 = objXmlDoc.getElementsByTagName(strRoot&"/YESTERDAY/COMMISSION")
	Set strSales4 = objXmlDoc.getElementsByTagName(strRoot&"/LASTWEEK/SALES")
	Set intSales4 = objXmlDoc.getElementsByTagName(strRoot&"/LASTWEEK/VOLUME")
	Set dblSales4 = objXmlDoc.getElementsByTagName(strRoot&"/LASTWEEK/COMMISSION")
	dArrTrueSales(0) = CDbl(strSales1.item(0).text)
	dArrTrueSales(1) = CDbl(strSales2.item(0).text)
	dArrTrueSales(2) = CDbl(strSales3.item(0).text)
	dArrTrueSales(3) = CDbl(strSales4.item(0).text)
	iArrTrueSales(0) = Int(intSales1.item(0).text)
	iArrTrueSales(1) = Int(intSales2.item(0).text)
	iArrTrueSales(2) = Int(intSales3.item(0).text)
	iArrTrueSales(3) = Int(intSales4.item(0).text)
	dArrTrueProfits(0) = CDbl(dblSales1.item(0).text)
	dArrTrueProfits(1) = CDbl(dblSales2.item(0).text)
	dArrTrueProfits(2) = CDbl(dblSales3.item(0).text)
	dArrTrueProfits(3) = CDbl(dblSales4.item(0).text)

	'Do predictions
	Set strSales1 = objXmlDoc.getElementsByTagName(strRoot&"/CURRENTMONTH/SALES")
	Set intSales1 = objXmlDoc.getElementsByTagName(strRoot&"/CURRENTMONTH/VOLUME")
	Set dblSales1 = objXmlDoc.getElementsByTagName(strRoot&"/CURRENTMONTH/COMMISSION")
	dArrTrueSales(4) = CDbl(strSales1.item(0).text)
	iArrTrueSales(4) = Int(intSales1.item(0).text)
	dArrTrueProfits(4) = CDbl(dblSales1.item(0).text)
	dArrTrueSales(5) = CalculateProjectedTarget(dArrTrueSales(4))
	iArrTrueSales(5) = Int(CalculateProjectedTarget(iArrTrueSales(4)))
	dArrTrueProfits(5) = CalculateProjectedTarget(dArrTrueProfits(4))

	Set strSales1 = nothing
	Set intSales1 = nothing
	Set dblSales1 = nothing
	Set strSales2 = nothing
	Set intSales2 = nothing
	Set dblSales2 = nothing
	Set strSales3 = nothing
	Set intSales3 = nothing
	Set dblSales3 = nothing
	Set strSales4 = nothing
	Set intSales4 = nothing
	Set dblSales4 = nothing
	Set objXmlDoc = nothing
	On Error Goto 0
End Sub

Private Sub updateShopInfo()
	Exit Sub
	
	Dim objXmlHttpCat, objXmlDoc, dblEuroToGo, dblDollarToGo, fso, f
	
	Set objXmlHttpCat= Server.CreateObject("MSXML2.ServerXMLHTTP")
	objXmlHttpCat.open "GET", "http://admin.drinkstuff.com/affiliate/shopinfo.asp" , False
	objXmlHttpCat.send ""
	Set objXmlDoc = Server.CreateObject("Microsoft.XMLDOM")
	strXML = objXmlHttpCat.responseXml.xml
	objXmlDoc.loadXML(strXML)
	Set objXmlHttpCat = nothing
		
	Set dblEuroToGo = objXmlDoc.getElementsByTagName("SHOPINFO/CURRENCIES/EURO")
	Set dblDollarToGo = objXmlDoc.getElementsByTagName("SHOPINFO/CURRENCIES/DOLLAR")
	
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	Set f = fso.CreateTextFile(Server.MapPath("/includes/shop/currency.asp"),True)
	f.writeLine("<" & "%")
	f.writeLine("Dim dblEuro, dblDollar")
	f.writeLine("dblEuro = " & dblEuroToGo.item(0).text)
	f.writeLine("dblDollar = " & dblDollarToGo.item(0).text)
	f.writeLine("dblEuro = 1")
	f.writeLine("dblDollar = 1")
	f.writeLine("%" & ">")
	f.close
	Set f = nothing
	
	Set objXmlDoc = nothing
End Sub

Function CalculateProjectedTarget(intCurrent)
	Dim intHoursPassed, intHoursTotal, dateFirst, dateDate

	dateDate = Now()
	dateFirst = CDate("1-" & MonthName(Month(dateDate), True) & "-" & Year(dateDate))
	intHoursTotal        = DateDiff("h", dateFirst, DateAdd("m", 1, dateFirst))
	intHoursPassed      = ((Day(dateDate)-1)* 24) + Hour(dateDate)
	
	CalculateProjectedTarget = CLng(CDbl(intCurrent)/CDbl(intHoursPassed)*intHoursTotal)
End Function

Sub sendCocktailsubmitEmail(strName, strEmail)
	Dim strBody, strSubject
	strSubject = "Cocktail:UK Drink Submission Confirmation"
	strBody= "<HTML><HEAD><TITLE>Cocktail:UK Drink Submission Confirmation</TITLE><LINK href=""http://www.cocktail.uk.com/style/mail.css"" type=""text/css"" rel=""stylesheet"" /></HEAD><BODY bgcolor=#ffffff><P>Hi,<BR>The cocktail you submitted, " & strName & ", has been viewed and has been added to the Cocktail : UK recipe list.<BR>Thank you for taking the time to submit the recipe, your input is highly appreciated and is what helps make the site.<BR><A href=""http://www.cocktail.uk.com"">http://www.cocktail.uk.com</A></BODY></HTML>"
	call SendEmail("theteam@cocktail.uk.com", strEmail, "", "", strSubject, strBody, True, "")
End Sub

Sub ReindexSite(strAction, strWhat, intNum, blnSilent)
	Dim objHTTP
	Exit Sub
	If NOT blnSilent Then
		Response.write ("&nbsp;&nbsp;<SPAN class=""linksin""><B>Re-indexing "& intNum & " " & strWhat & "</B></SPAN><BR>")
	End If
	Set objHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	objHTTP.open "GET", "http://www.cocktailuk.com/admin/sitesearch/default.asp?action="&strAction , False, "lee", "Smetsy#190"
	objHTTP.send ""
	Set objHTTP = Nothing
End Sub

Sub getDateShopLastUpdated(cn, rs, strWhen, dteLastUpdated)
	Dim intNumDays
	rs.Open "SELECT dteshopupdated from dsshopupdate", cn, 0, 3
	If NOT rs.EOF Then
		dteLastUpdated = FormatDateTime(rs(0),2)
		intNumDays = DateDiff("d", dteLastUpdated, Now())
		Select Case intNumDays
			Case 0 strWhen = "today"
			Case 1 strWhen = "yesterday"
			Case Else strWhen = intNumDays & " days ago"
		End Select 
	Else
		dteLastUpdated = ""
	End If
	rs.Close
End Sub
%>