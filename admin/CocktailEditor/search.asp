<%Option Explicit%>
<!--#include file="cdbviewer.inc"-->
<!--#include virtual="/includes/variables.asp"-->
<%
CONST ID			= 0
CONST STATUS		= 1
CONST TYPEID		= 2
CONST REINDEX		= 3
CONST SERVES		= 4
CONST BASED			= 5
CONST RATE			= 6
CONST USERS			= 7
CONST ACCESSED		= 8
CONST NAME			= 9
CONST INGREDIENTS	= 10
CONST USER			= 11
CONST DESCRIPTION	= 12

Class CCocktailView
	Public m_objDBViewer

	Public Sub Class_Initialize
		Set m_objDBViewer 				= New CDBViewer
		m_objDBViewer.m_strTableName 	= "Cocktail"
		m_objDBViewer.SetParentObject( me )

		' CHANGE: Add additional fields here

		Call m_objDBViewer.AddField(ID, "ID", "ID", tID OR tNUMERIC, "", fEQUAL, False, DEFAULT)
		Call m_objDBViewer.AddField(STATUS, "Status", "Status", tNUMERIC, "", fEQUAL, False, DEFAULT)
		Call m_objDBViewer.AddField(TYPEID, "Type", "Type", tNUMERIC, "", fEQUAL, False, DEFAULT)
		Call m_objDBViewer.AddField(REINDEX, "Reindex", "Reindex", tNUMERIC, "", fEQUAL, False, DEFAULT)
		Call m_objDBViewer.AddField(SERVES, "Serves", "Serves", tNUMERIC, "", fEQUAL, False, DEFAULT)
		Call m_objDBViewer.AddField(BASED, "Based", "Based", tNUMERIC, "", fEQUAL, False, DEFAULT)
		Call m_objDBViewer.AddField(RATE, "Rate", "Rate", tNUMERIC, "", fEQUAL, False, DEFAULT)
		Call m_objDBViewer.AddField(USERS, "Users", "Users", tNUMERIC, "", fEQUAL, False, DEFAULT)
		Call m_objDBViewer.AddField(ACCESSED, "Accessed", "Accessed", tNUMERIC, "", fEQUAL, False, DEFAULT)
		Call m_objDBViewer.AddField(NAME, "Name", "Name", tVARCHAR, "", fLIKE, False, DEFAULT)
		Call m_objDBViewer.AddField(USER, "User", "usr", tVARCHAR, "", fLIKE, False, DEFAULT)
		Call m_objDBViewer.AddField(DESCRIPTION, "Description", "Description", tTEXT, "", fLIKE, False, DEFAULT)
	End Sub

	Public Sub DisplayFilter(intEntryID)
		' CHANGE: Add any special case code calls here...
		Select Case intEntryID
			Case STATUS			Call m_objDBViewer.DisplayStandardDropdownFilter(Array(0,1,2), Array("Pending","Live","Deleted"), intEntryID)
			Case TYPEID			Call m_objDBViewer.DisplayStandardDropdownFilter(Array(1,2,5,6,9,10), Array("Cocktail","Shooter","Cocktail (NA)","Shooter (NA)","Cocktail XXX","Shooter XXX"), intEntryID)
			Case REINDEX		Call m_objDBViewer.DisplayStandardDropdownFilter(Array(0,1), Array("No","Yes"), intEntryID)
			Case Else			Call m_objDBViewer.DisplayStandardFilter(intEntryID)
		End Select
	End Sub

	Public Function LinkEdit(intEntryID)
		Select Case intEntryID
			Case ID LinkEdit = "<a target=""edit"" href=""edit.asp?id=" & m_objDBViewer.m_aryData(ID) & """>"
		Case Else
				LinkEdit = ""
		End Select
	End Function

	Public Function MapData(intEntryID)
		Select Case intEntryID
			Case STATUS		MapData = m_objDBViewer.MapValue(Array(0,1,2), Array("Pending","Live","Deleted"), intEntryID)
			Case TYPEID		MapData = m_objDBViewer.MapValue(Array(1,2,5,6,9,10), Array("Cocktail","Shooter","Cocktail (NA)", "Shooter (NA)","Cocktail XXX","Shooter XXX"), intEntryID)
			Case REINDEX	MapData = m_objDBViewer.MapValue(Array(0,1), Array("No","Yes"), intEntryID)
			Case Else		MapData = m_objDBViewer.m_aryData(intEntryID)
		End Select
	End Function

End Class

Dim objDBViewer
Set objDBViewer		= New CCocktailView
objDBViewer.m_objDBViewer.m_strDSN = strDB
objDBViewer.m_objDBViewer.m_strApplicationName = "Cocktail"

Call objDBViewer.m_objDBViewer.DisplayTop()
Call objDBViewer.m_objDBViewer.DisplayBottom()
Set objDBViewer		= Nothing
%>