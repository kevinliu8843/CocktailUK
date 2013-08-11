<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/admin_functions.asp" -->
<!--#include virtual="/includes/classes/clsSiteMap.asp" -->
<%
Server.ScriptTimeout = 10 * 60

Dim objSiteMap, fso, objFile, aryExtraPages, aryExtraPriority, aryExtraFrequency, i, strPriority

strUnSecureUrl = "http://www.cocktail.uk.com/"

Set objSiteMap = New clsSiteMap

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.open strDB
Set cn2= Server.CreateObject("ADODB.Connection")
Set rs2= Server.CreateObject("ADODB.Recordset")
cn2.open strDB

'Call setupCategories(NULL) 

'Generate pretty URL's whils we are here too...
Call CreatePrettyURLFiles(cn, rs)

'Add surplus pages to site map
Call AddURL("default.asp", 	"1",   "daily")

'Add the surplus pages
For i=0 To UBound(aryExtraPages)
	Call objSiteMap.AddSiteMapURL(strUnSecureUrl & aryExtraPages(i), "", aryExtraFrequency(i), aryExtraPriority(i))
Next 

'Write out all recipes
rs.open "SELECT name, type FROM cocktail WHERE status=1 ORDER BY accessed", cn
If NOT rs.EOF Then
	aryRecipes = rs.GetRows()
Else
	ReDim aryRecipes(-1, -1)
End If
rs.close

For i=0 To UBound(aryRecipes, 2)
	If Int(aryRecipes(1, i)) AND 1 Then
		strType = "Cocktail"
	ElseIf Int(aryRecipes(1, i)) AND 2 Then
		strType = "Shooter"
	End If
	Call objSiteMap.AddSiteMapURL(strUnSecureUrl & strType & "-Recipe/"&GeneratePrettyURL(aryRecipes(0, i))&".htm", "", "weekly", "0.6")
Next

'Write out all parent categories
rs.open "SELECT ID, parentID, url FROM DScategory WHERE hidden=0 AND parentID=0 ORDER by catorder", cn
While NOT rs.EOF
	Call WriteCategories(rs("ID"), rs("parentID"), 10, "shop/products", rs("url"), cn2, rs2)
	rs.MoveNext
Wend
rs.close

'Write out all products 
strSQL = "SELECT DSproduct.ID, name FROM DSproduct "
strSQL = strSQL & " WHERE DSproduct.status=1 "
strSQL = strSQL & " ORDER BY DSproduct.ID"

rs.open strSQL, cn
While NOT rs.EOF
	Call objSiteMap.AddSiteMapURL(strUnSecureUrl & "shop/products/" & GeneratePrettyURL(rs("name")) & ".htm", "", "weekly", 0.5)
	rs.MoveNext
Wend
rs.close

cn.close
Set cn = Nothing
Set rs = Nothing
cn2.close
Set cn2= Nothing
Set rs2= Nothing

Set objSiteMap = Nothing
Response.end

Sub AddURL(strFile, strPriority, strFrequency)
	If NOT IsArray(aryExtraPages) Then
		ReDim aryExtraPages(0)
	Else
		ReDim Preserve aryExtraPages(UBound(aryExtraPages)+1)
	End If
	
	If NOT IsArray(aryExtraPriority) Then
		ReDim aryExtraPriority(0)
	Else
		ReDim Preserve aryExtraPriority(UBound(aryExtraPriority)+1) 
	End If
	
	If NOT IsArray(aryExtraFrequency) Then
		ReDim aryExtraFrequency(0)
	Else
		ReDim Preserve aryExtraFrequency(UBound(aryExtraFrequency)+1)
	End If 
	
	aryExtraPages(UBound(aryExtraPages)) = strFile
	aryExtraPriority(UBound(aryExtraPriority)) = strPriority
	aryExtraFrequency(UBound(aryExtraFrequency)) = strFrequency
End Sub

Sub WriteCategories(catID, parentID, pagesize, strFolder, strFile, cn2, rs2)
	Dim aryCats, i, intPages, intProds
	
	Call objSiteMap.AddSiteMapURL(strUnSecureUrl & strFolder & "/" & strFile & ".asp", "", "weekly", "0.6")
	
	If catID <> 0 Then
		rs2.open "SELECT count(*) as numprods FROM DSprodcat WHERE catID=" & catID, cn2
		intProds = rs2("numprods")
		rs2.close
	Else
		intProds = 0
	End If
	
	'Do category pages and a view all if more than 1 page
	If IsNull(pagesize) OR pagesize = "" OR pagesize = 0 Then
		pagesize = 10
	End If
	
	If pagesize > 0 AND intProds > 0 Then
		intPages = Int(intProds / pagesize) + 1
		
		For i=2 To intPages
			Call objSiteMap.AddSiteMapURL(strUnSecureUrl & strFolder & "/" & strFile & ".asp?page=" & i, "", "weekly", "0.6")
		Next
	End If
	
	rs2.open "SELECT ID, parentID, url FROM DScategory WHERE hidden=0 AND parentID=" & catID & "  ORDER by catorder", cn2
	If NOT rs2.EOF Then
		aryCats = rs2.GetRows()
	Else
		ReDim aryCats(-1, -1)
	End If 
	rs2.close
	
	For i=0 To UBound(aryCats, 2)
		Call WriteCategories(aryCats(0, i), aryCats(1, i), 10, strFolder, aryCats(2, i), cn2, rs2)
	Next
End Sub
%>