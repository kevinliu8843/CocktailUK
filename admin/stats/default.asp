<%
Option Explicit
strtitle="Cocktail : UK Statistics"
%>
<!--#include virtual="/includes/charting/cbarchart.asp" -->
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/admin_functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<H2>Cocktail : UK Statistics</H2>
<DIV align="center"> 
<A href="default.asp">Usage stats</A> | <A href="default.asp?type=shop">Shop stats</A><TABLE COLS="2" width="95%">
<%
Dim cn, objBarChart, dateNow, dateDate, newDateAdd6, newDateLess6, intProj
Dim dblAverage1, dblAverage2, dblDistance, dblGradient, iStep, blnShop

set cn	= Server.CreateObject("ADODB.Connection")
Set rs	= Server.CreateObject("ADODB.RecordSet")
cn.Open strDBMod

blnShop = (Request("type") = "shop")

dateNow = Now()
If Request("dtDate") = "" Then
	dateDate = dateNow
Else
	dateDate = CDate(Request("dtDate"))
End If

Set objBarChart = New CBarChart

If NOT blnShop then

	Response.Write("<TR><TD colspan=2>")
	objBarChart.SetChartHeight(150)
	objBarChart.SetBarWidth(22)
	objBarChart.SetBarGap(8)
	Call OutputChart("visitors", "Visitors/Month")
	Response.Write("</TD></TR><TR><TD>")

	Call objBarChart.Reset()
	objBarChart.SetChartHeight(150)
	objBarChart.SetBarWidth(8)
	objBarChart.SetBarGap(3)
	Call OutputChart("cocktails", "Drinks/Month")
	Response.Write("</TD><TD>")
	
	objBarChart.SetChartHeight(150)
	objBarChart.SetBarWidth(8)
	objBarChart.SetBarGap(3)
	Call OutputChart("ajsearches", "Searches/Month")
	Response.Write("</TD></TR>")
Else
	Response.Write("<TR><TD>")

	Call objBarChart.Reset()
	objBarChart.SetChartHeight(150)
	objBarChart.SetBarWidth(8)
	objBarChart.SetBarGap(3)
	Call OutputChart("products", "Views/Month")
	Response.Write("</TD><TD>")
	
	Call objBarChart.Reset()
	objBarChart.SetChartHeight(150)
	objBarChart.SetBarWidth(8)
	objBarChart.SetBarGap(3)
	Call OutputChart("categories", "Views/Month")
	Response.Write("</TD></TR><TR><TD>")
	
	Call objBarChart.Reset()
	objBarChart.SetChartHeight(150)
	objBarChart.SetBarWidth(8)
	objBarChart.SetBarGap(3)
	Call OutputChart("baskets", "Filled per month")
	Response.Write("</TD><TD>")
	
	Call objBarChart.Reset()
	objBarChart.SetChartHeight(150)
	objBarChart.SetBarWidth(8)
	objBarChart.SetBarGap(3)
	Call OutputChart("equipment", "Referals/Month")
	Response.Write("</TD></TR>")
End If

Set objBarChart = Nothing

newDateless6	= DateAdd("m", -6, dateDate)
newDateadd6		= DateAdd("m", 6, dateDate)
%>
</TABLE>
<P><A HREF="default.asp?dtDate=<%=Day(newDateless6)%>/<%=Month(newDateless6)%>/<%=Year(newDateless6)%>">&lt; &lt; Previous period</A> | 
<A HREF="default.asp?dtDate=<%=Day(newDateadd6)%>/<%=Month(newDateadd6)%>/<%=Year(newDateadd6)%>">Next period &gt; &gt;</A></P>
</DIV>

<P align="center"><IMG border="0" src="../../images/side_menus/hlinebase.gif"></P>

<BLOCKQUOTE>

<H4>Database statistics</H4>
	<%
	cn.close
	cn.open strDB
	strSQL = "SELECT count(*) from cocktail WHERE Status=1 And type=(type | 1)"
	rs.open strSQL, cn
	%>
	<p>Total number of cocktails in database: <FONT color=red><%=FormatNumber(rs(0), 0)%></font>
	<%
	rs.close
	strSQL = "SELECT count(*) from cocktail WHERE Status=1 And type=(type | 2)"
	rs.open strSQL, cn
	%>
	<p>Total number of shooters in database: <FONT color=red><%=FormatNumber(rs(0), 0)%></font>
	<%
	rs.close
	strSQL = "SELECT count(*) from usr"
	rs.open strSQL, cn
	%>
	<p>Total number of registered users in database: <FONT color=red><%=FormatNumber(rs(0), 0)%></font></BLOCKQUOTE>

<P align="center"><IMG border="0" src="../../images/side_menus/hlinebase.gif"></P>
<%
cn.Close
set cn			= Nothing
Set rs			= Nothing
%><!--#include virtual="/includes/footer.asp" -->