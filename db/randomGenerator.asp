<%
Option Explicit
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
Dim cn, bLoggedIn, blnAbleToMake, strRecipe, strList
Dim strIngredients
Dim strSpiritName, intPrevID, i
Dim aryNeeded, aryHowMany, aryLowerLimit, aryUpperLimit, aryIngredients

ReDim aryNeeded(g_intNumIngredientTypes)
ReDim aryHowMany(g_intNumIngredientTypes)
ReDim aryLowerLimit(g_intNumIngredientTypes)
ReDim aryUpperLimit(g_intNumIngredientTypes)
ReDim aryIngredients(g_intNumIngredientTypes)


strTitle = "Random Cocktail Generator"

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.RecordSet")
cn.Open strDB

bLoggedIn = Session("logged")
blnAbleToMake = True
%>
<!--#include virtual="/includes/header.asp" -->
<%
If bLoggedIn Then
	'Check for an ingredients list
	blnAbleToMake = False
	strSQL = "SELECT Top 1 IngredientID FROM UsrIng WHERE memID=" & Session("ID")
	rs.Open strSQL, cn, 0, 3

	If Not rs.EOF Then
		blnAbleToMake = True
	End If
	rs.Close
End If

strRecipe = ""

IF blnAbleToMake Then
	'Randomly select the number of the parts of this cocktail

	aryLowerLimit(SPIRIT)		= 1
	aryLowerLimit(LIQUOR)		= 0
	aryLowerLimit(JUICE)		= 0
	aryLowerLimit(MIXER)		= 1
	aryLowerLimit(WINE)			= 0
	aryLowerLimit(FRUIT)		= 0
	aryLowerLimit(FLAVOURINGS)	= 0
	aryLowerLimit(SYRUPS)		= 0

	aryUpperLimit(SPIRIT)		= 3
	aryUpperLimit(LIQUOR)		= 2
	aryUpperLimit(JUICE)		= 1
	aryUpperLimit(MIXER)		= 2
	aryUpperLimit(WINE)			= 1
	aryUpperLimit(FRUIT)		= 1
	aryUpperLimit(FLAVOURINGS)	= 1
	aryUpperLimit(SYRUPS)		= 1

	For i=0 to g_intNumIngredientTypes
		aryNeeded(i)		= randomise ( aryUpperLimit(i), aryLowerLimit(i) )
	Next

	If bLoggedIn Then
		Call SplitUpIntoCategories(aryIngredients)
		
		For i=0 to g_intNumIngredientTypes
			aryHowMany(i) = howMany(aryIngredients(i))
			aryUpperLimit(i) = min( aryHowMany(i), aryUpperLimit(i))

			If aryHowMany(i) > 0 Then
				aryNeeded(i) = randomise ( aryUpperLimit(i), aryLowerLimit(i) )
			Else
			aryNeeded(i) = 0
			End If
		Next
	End If

	'----------------------------------------------------------------------

	intPrevID = ""
	For i=0 to g_intNumIngredientTypes
		Call generateIdList( g_aryIngredientTypeID(i), aryNeeded(i), aryIngredients(i))
	Next
	'--------------------------------------------------------------

	cn.Close
	Set cn = Nothing
	Set rs = Nothing
%>
  <h2>The Cocktail Generator</h2>

  <TABLE border="0" cellpadding="0" cellspacing="5" style="border-collapse: collapse" bordercolor="#111111">
    <TR>
      <TD valign="top" colspan="3">
      <P align="justify">This is a randomly <I>generated </I>cocktail.
        i.e. has random number/selection of ingredients. It may be nice or
        completely the opposite! Select the appropriate amount of each
        ingredient to suit your taste. <BR>
      Don't blame us for the headache after!</TD>
    </TR>
    <TR>
      <TD valign="top">
			<TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
			  <TBODY>
			  <TR>
				<TD class=arrowblock align=left width="1%" nowrap>
                <IMG 
				  height=15 
				  src="/images/pixel.gif" 
				  width=15 border=0></TD>
				<TD class=baselightred width="99%"><B 
				  class=contentHeader>&nbsp;DIRECTIONS</B></TD>
				</TR></TBODY></TABLE>
			Make as appropriate.
        </TD>
      <TD valign="top">
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <TR>
                <TD class=arrowblock align=left width="1%" nowrap>
                <IMG 
				  height=15 
				  src="/images/pixel.gif" 
				  width=15 border=0></TD>
                <TD class=baselightred width="99%"><B 
                  class=contentHeader>&nbsp;INGREDIENTS</B></TD>
                </TR></TBODY></TABLE>
        <nobr><%=strRecipe%>Ice cubes (as desired)</nobr>
      </TD>
      <%Randomize%>
      <TD valign="top" rowspan="3"> <IMG src="/images/cocktails/default/cocktail<%=Int((6 - 1) * Rnd + 1)%>.jpg">
      </TD>
    </TR>
    <TR>
      <TD valign="top" colspan="2">
        <b>Serves:<BR>
        </b>
        However many you want it to! <BR>
        <BR>
        <img border="0" src="/images/home_mod_bullet.gif" width="6" height="6"><B>
        <A href="randomGenerator.asp">MAKE ANOTHER DRINK!!!</A></B><p>&nbsp;</TD>
    </TR>
  </TABLE>

<%
Else	'no ingredients list%>
<H2>Random Cocktail Generator</H2>
<p>Whoops...you don't appear to have enough ingredients saved in your web-bar<BR>
<a href="/db/member/selectIngredients.asp">Modify your web-bar now</a>

<%End If%><%Call DrawSearchCocktailArea()%><!--#include virtual="/includes/footer.asp" --><%
Function generateIdList( intType, intCat, strList)
	Dim upperbound, lowerbound, i, j, intID, intPos, strTemp

	If bLoggedIn Then
		If Len(strList) > 0 Then
			strSQL = "SELECT DISTINCT ID, name FROM Ingredients WHERE Status=1 And Type="&intType&" AND ID IN (SELECT IngredientID FROM UsrIng WHERE memID="&Session("ID")&")"
		Else
			strSQL = "SELECT DISTINCT ID, name FROM Ingredients WHERE Status=1 And Type = " & intType
		End If
	Else
		strSQL = "SELECT DISTINCT ID, name FROM Ingredients WHERE Status=1 And Type = " & intType
	End If
	rs.Open strSQL, cn, 3, 3

	upperbound = rs.recordcount
	lowerbound = 1

	For i=1 to intCat
		intID = randomise ( upperbound, lowerbound )

		'randomise up to 3 times
		If rs.recordcount > 1 Then
			For j=1 to 3
				If intID = intPrevID Then
					intID = randomise ( upperbound, lowerbound )
				End If
			Next

			If intID <> 0 Then rs.absolutePosition = intID
		End If

		intPrevID = intID
		strRecipe = strRecipe & "<A HREF=""ingredient_description.asp?id="&rs("ID") & """>" & capitalise( rs("name") ) & "</A><BR>"
	Next

	rs.Close
End Function

'------------------------------------------------------------------------------------------

Function randomise( upperbound, lowerbound )
	'Random ID generator
	'Returns an ID to be checked
	Randomize
	If upperbound = lowerbound then
		randomise = lowerbound
	else
		randomise = Int((upperbound - lowerbound ) * Rnd + lowerbound)
	end if
End Function

'------------------------------------------------------------------------------------------

Function howMany( strIng )
	Dim MyArray
	'Returns an integer to show how many ingredients the user has saved for that category
	MyArray = Split( strIng, ",", -1, 1)
	howMany = UBound(MyArray)
	If howMany = 0 AND Len( strIng ) > 0 Then
		howMany = 1
	End If
End Function

Sub SplitUpIntoCategories(aryIngredients)
	' Takes one list and splits into the four components
	Dim strList, intPos, i

	intPos = InstrRev(strIngredients, "," )
	If intPos > 0 Then
		strList = Left(strIngredients, intPos-1)
	Else
		strList = strIngredients
	End If
	
	For i=0 to g_intNumIngredientTypes
		aryIngredients(i) = ""
	Next

	strSQL = "SELECT ID, Type FROM Ingredients WHERE Status=1 AND ID IN (SELECT IngredientID FROM UsrIng WHERE memID="&Session("ID")&")"
	rs.Open strSQL, cn, 0, 3
	Do While Not rs.EOF
		For i=0 to g_intNumIngredientTypes
			If Int(rs("Type")) = Int(g_aryIngredientTypeID(i)) Then
				aryIngredients(i) = aryIngredients(i) & rs("ID") & ","
				Exit For
			End If
		Next
		rs.MoveNext
	Loop
	rs.Close
End Sub

%>