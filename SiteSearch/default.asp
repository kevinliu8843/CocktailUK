<%
Option Explicit
Server.ScriptTimeout = 30
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%
Dim timePageLoadStart, strImgSrc, blnBGColour, strBGColour 
Dim objSearch, intSearchType, bHasMore, intPage
Dim blnGeneral, blnForums, blnDrinks, blnShop

Const SEARCH			= "s"
Const SEARCHTYPE		= "t"
Const SEARCHORDER		= "o"

CONST MAXPAGES			= 16

blnBGColour = True
strTitle = "Search"

timePageLoadStart = Timer()

Set rs			= Server.CreateObject("ADODB.RecordSet")
Set objSearch	= New CSearch
objSearch.m_intPageSize = 10

blnGeneral = False
blnForums  = False
blnDrinks  = False
blnShop    = False
If Request("site_pages") = "ON" Then
	Call objSearch.AddFilter(SEARCH_TYPE_GENERAL)
	blnGeneral = True
End If
If Request("forum_posts") = "ON" Then
	Call objSearch.AddFilter(SEARCH_TYPE_FORUMS)
	blnForums = True
End If
If Request("drink_recipes") = "ON" Then
	Call objSearch.AddFilter(SEARCH_TYPE_DRINK)
	blnDrinks = True
End If
If Request("theshop") = "ON" Then
	Call objSearch.AddFilter(SEARCH_TYPE_SHOP)
	blnShop = True
End If

If NOT (blnDrinks OR blnForums OR blnDrinks OR blnShop) Then
	blnGeneral = True
	blnForums  = True
	blnDrinks  = True
	blnShop    = True
End If

intPage = (Request("pg"))
If intPage <> "" Then
	objSearch.m_intCurrentPage = Int(intPage)
End If
If intPage = "" Then
	intPage = 1
Else
	intPage = Int(intPage)
End If

objSearch.m_strSearchString = Replace(Request(SEARCH), "'", "")
If Trim(Request("searchField")) <> "" Then
	objSearch.m_strSearchString = Request("searchField")
	Call objSearch.AddFilter(SEARCH_TYPE_GENERAL)
	Call objSearch.AddFilter(SEARCH_TYPE_FORUMS)
	Call objSearch.AddFilter(SEARCH_TYPE_DRINK)
	Call objSearch.AddFilter(SEARCH_TYPE_SHOP)
	blnGeneral = True
	blnForums  = True
	blnDrinks  = True
	blnShop    = True
Else
	'blnGeneral = True
	'blnForums  = True
	'blnDrinks  = True
	'blnShop    = True
End If

If Request(SEARCHTYPE) <> "" Then
	objSearch.m_intSearchType = Int(Request(SEARCHTYPE))
End If

If Request(SEARCHORDER) <> "" Then
	objSearch.m_intSortOrder = Int(Request(SEARCHORDER))
End IF
%>

<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/search/csearch.inc" -->
<!--#include virtual="/includes/search/cstopword.inc" -->
<SCRIPT Language="JavaScript">
function addFilter(){
	if (!((document.searchform.site_pages.checked) || (document.searchform.forum_posts.checked) || (document.searchform.drink_recipes.checked) || (document.searchform.theshop.checked))){
		alert("Please select some search criteria...")
		document.searchform.site_pages.checked = true
	}
}
</SCRIPT>

<BR>
<form method="GET" action="default.asp" style="text-align: center" name="searchform">

<input type="hidden" name="update" value="1">
<input type="hidden" name="pg" value="1">
<input type="hidden" name="<%=SEARCHORDER%>" value="<%=objSearch.m_intSortOrder%>">
	    <TABLE border="0" cellpadding="0" cellspacing="0" style="border-style:solid; border-width:1; border-collapse: collapse" bordercolor="#612B83" id="AutoNumber6" height="150" width="0">
          <TR>
            <TD height="42" background="../images/main_menus/searchcocktailuk.gif">
	    &nbsp;</TD>
          </TR>
          <TR>
            <TD background="../images/grad_write_purple.gif">
<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
  <TR>
    <TD valign="top">&nbsp;</TD>
    <TD valign="top">

<P align="center" style="margin-bottom: 20px;">
<B>Search: </B> 
<input type="text" class="shopoption" name="<%=SEARCH%>" value="<%=objSearch.m_strSearchString%>" size="55"><input type="image" src="../images/template/button_search_go.gif" align="absmiddle" id="go" alt="Search" title="Search"> 
<DIV align="center">
  <CENTER>
  <TABLE border="0" cellpadding="05" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse">
    <TR>
      <TD width="33%" nowrap><INPUT type="checkbox" onclick="addFilter()" name="site_pages" ID="site_pages" value="ON" <%If blnGeneral Then%>checked<%End If%>><label for="site_pages"><IMG border="0" src="../images/sitesearch/information.gif" align="middle" width="31" height="31">Site Pages</label></TD>
      <TD width="33%" nowrap><INPUT type="checkbox" onclick="addFilter()" name="forum_posts" ID="forum_posts" value="ON" <%If blnForums Then%>checked<%End If%>><label for="forum_posts"><IMG border="0" src="../images/sitesearch/forum.gif" align="middle" width="31" height="31">Forums</label></TD>
      <TD width="34%" nowrap><INPUT type="checkbox" onclick="addFilter()" name="drink_recipes" ID="drink_recipes" value="ON" <%If blnDrinks Then%>checked<%End If%>><label for="drink_recipes"><IMG border="0" src="../images/sitesearch/cocktail.gif" align="middle">Drinks</label></TD>
      <TD width="34%" nowrap><INPUT type="checkbox" onclick="addFilter()" name="theshop" ID="theshop" value="ON" <%If blnShop Then%>checked<%End If%>><label for="theshop"><IMG border="0" src="../images/sitesearch/shop.gif" align="middle">Shop</label></TD>
    </TR>
    <TR>
      <TD width="200%" nowrap colspan="4">
      <P align="center"><B>OR search for a drink <A href="../db/search/searchByAlphabet.asp">alphabetically</A> OR by
      <A href="../db/search/searchByIngredient.asp">ingredient</A></B></TD>
    </TR>
  </TABLE>
  </CENTER>
</DIV>
    </TD>
    <TD valign="top">&nbsp;</TD>
  </TR>
</TABLE>
            </TD>
          </TR>
        </TABLE>
    <br>
</form>
<%
If intPage <> "" OR Request("searchField") <> "" Then
	Call objSearch.DoSearch()

	If objSearch.HasResults() Then
		Response.write("<CENTER><H4>Results for <FONT color=#888888>"""&Trim(objSearch.m_strSearchString)&"""</FONT></H4></CENTER>")
	
		bHasMore = objSearch.GetFirst()
		blnBGColour  = False
		Do While bHasMore
			blnBGColour  = NOT blnBGColour 

			If blnBGColour Then
				strBGColour = "#EFEFEF"
			Else
				strBGColour = "#FFFFFF"
			End If

			Response.Write("<TABLE border=""0"" cellpadding=""5"" cellspacing=""0"" align=""center"" width=""100%"" bgcolor="""&strBGColour&""">")
			Response.Write("<TR>")
			Response.Write("<TD valign=""top"" align=""left"" width=""35""><B><FONT size=""3"">"&objSearch.m_intLineNumber&" .</FONT></B></TD>")

			Select Case objSearch.m_intTypeID
				Case SEARCH_TYPE_GENERAL				strImgSrc = "/images/sitesearch/information.gif"
				Case SEARCH_TYPE_FORUMS					strImgSrc = "/images/sitesearch/forum.gif"
				Case SEARCH_TYPE_DRINK					strImgSrc = "/images/sitesearch/cocktail.gif"
				Case SEARCH_TYPE_SHOP					strImgSrc = "/images/sitesearch/shop.gif"
			End Select

			Response.Write("<TD valign=""top"" align=""center"" width=""20""><A href="""&objSearch.m_strURL&"""><IMG SRC=""" & strImgSrc & """ height=""31"" border=""0"" align=""center""></A></TD>")
			Response.Write("<TD valign=""top"" align=""left""><A href=""" & objSearch.m_strURL & """><FONT size=""3""><B>" & ReplaceStuffBack(objSearch.m_strTitle) & "</B></FONT></A><BR>")

			strTopTitle = ""
			strTitleOut = ""
			call displayPageLocation(objSearch.m_strTitle , strTitleOut, strTopTitle, objSearch.m_strURL, "color: darksilver")

			Response.Write("<SPAN style=""color: darksilver; font-size: 7pt;"">"&strTitleOut&"<BR>")
			Response.Write("Relevancy :</SPAN>")

			Call displayRatingGraphOnly(objSearch.m_fltHitScore)

			Response.Write("</TD>")
			Response.Write("</TR>")
			Response.Write("</TABLE>")

			bHasMore = objSearch.GetNext()
		Loop

		If objSearch.m_intTotalPages > 1 Then
			Response.Write("<CENTER><BR>" & GetPageControls() &"</CENTER>" )
		End If
 
		Call ShopSearch(objSearch.m_strSearchString)
	Else
		Call ShopSearch(objSearch.m_strSearchString)
		Call objSearch.ShowError()
		' Response.write("<BR><CENTER><IFRAME frameborder=""0"" scrolling=""no"" id=""s0"" name=""s0"" align=absmiddle border=0 height=56 width=340 src=""/db/search/ask_jeeves.asp""></IFRAME></CENTER>")
	End If
End If
 
Set objSearch		= Nothing

Response.write("<P>")

Function GetPageControls()
	Dim i, strSite, strForum, strDrink, intStart, intFinish 

	If blnGeneral Then
		strSite = strSite & "&site_pages=ON"
	End If
	If blnForums Then
		strSite = strSite & "&forum_posts=ON"
	End If
	If blnDrinks Then
		strSite = strSite & "&drink_recipes=ON"
	End If
	If blnShop Then
		strSite = strSite & "&theshop=ON"
	End If

	If Int(objSearch.m_intCurrentPage) > 1 Then
		GetPageControls = GetPageControls & "<a href=""" & GetHREF(objSearch.m_intCurrentPage-1, objSearch.m_strSearchString, objSearch.m_intSearchType, objSearch.m_intSortOrder) & strSite & """>&lt;&lt; Previous</a>&nbsp;"
	Else
		GetPageControls = GetPageControls & "&laquo; Previous&nbsp;"
	End If

	intStart	= Max(1, objSearch.m_intCurrentPage-MAXPAGES/2)
	intFinish	= Min(intStart+MAXPAGES-1, objSearch.m_intTotalPages)

	If intFinish-intStart <= MAXPAGES Then
		intStart = intFinish - MAXPAGES + 1
		If intStart<1 then
			intStart=1
		End If
	End If

	For i=intStart To intFinish
		If Int(i) = Int(objSearch.m_intCurrentPage) Then
			GetPageControls = GetPageControls & "<b>[" & i & "]</b>&nbsp;"
		Else
			GetPageControls = GetPageControls & "<a href=""" & GetHREF(i, objSearch.m_strSearchString, objSearch.m_intSearchType, objSearch.m_intSortOrder) & strSite & """>" & i & "</a>&nbsp;"
		End If
	Next

	If Int(objSearch.m_intCurrentPage) < Int(objSearch.m_intTotalPages) Then
		GetPageControls = GetPageControls & "<a href=""" & GetHREF(objSearch.m_intCurrentPage+1, objSearch.m_strSearchString, objSearch.m_intSearchType, objSearch.m_intSortOrder) & strSite & """>Next &gt;&gt;</a>"
	Else
		GetPageControls = GetPageControls & "Next &raquo;"
	End If
	
	GetPageControls = GetPageControls & "<BR>&nbsp;"
End Function

Function GetHREF(intPage, strSearch, intSearchType, intSortOrder)
	GetHREF = "default.asp?pg=" & intPage 
	GetHREF = GetHREF & "&" & SEARCH & "=" & Server.URLEncode(Trim(strSearch))
	GetHREF = GetHREF & "&" & SEARCHTYPE & "=" & intSearchType
	GetHREF = GetHREF & "&" & SEARCHORDER & "=" & intSortOrder
End Function

Sub WriteHeader(strHeading, intHeading, intDefaultDir)
	If objSearch.m_intSortOrder And intHeading Then
		Response.Write("<b>")
	Else
		Response.Write("<a href=""" & GetHREF(objSearch.m_intCurrentPage, objSearch.m_strSearchString, objSearch.m_intSearchType, intHeading OR intDefaultDir) & """>")
	End If

	Response.Write(strHeading)
	If objSearch.m_intSortOrder And intHeading Then
		Response.Write("</b>")
	Else
		Response.Write("</a>")
	End If
	If objSearch.m_intSortOrder And intHeading Then
		If objSearch.m_intSortOrder And ASCENDING Then
			Response.Write("&nbsp;<a href=""" & GetHREF(objSearch.m_intCurrentPage, objSearch.m_strSearchString, objSearch.m_intSearchType, intHeading OR DESCENDING) & """>down</a>")
		Else
			Response.Write("&nbsp;<a href=""" & GetHREF(objSearch.m_intCurrentPage, objSearch.m_strSearchString, objSearch.m_intSearchType, intHeading OR ASCENDING) & """>up</a>")
		End If
	End If
End Sub

Sub ShopSearch(strSearch)
Exit Sub
	If strSearch <> "" Then
		%>
		<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1"><TR><TD class="shopheadertitle"><H3><a href="/shop/products/search.asp?search=<%=strSearch%>"><span style="text-decoration: none"><font color="#612B83">Click here to search for <%=strSearch%> in our shop</font></span></a></H3></TD></TR></TABLE>
		<%
	End If
End Sub
%><!--#include virtual="/includes/footer.asp" -->