<%
Option Explicit 
strTitle = "Party Planner"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<H2>Party planner</H2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
   <tr>
     <td width="100%">
<%If Request("number") = "" Then%>
	<FORM method="POST" action="party_planner.asp">
	<INPUT type="hidden" name="party" value="1">
	<P>This little tool should give you some idea of what things you will need to do
	and buy to prepare for a party.</P>
	<P>There will be 
	  <SELECT size="1" name="number">
	  <OPTION value="20"selected>&lt; 20</OPTION>
	  <OPTION value="30">20 - 30</OPTION>
	  <OPTION value="40">30 - 40</OPTION>
	  <OPTION value="50">40 - 50</OPTION>
	  <OPTION value="60">50 - 60</OPTION>
	  <OPTION value="70">60 - 70</OPTION>
	  <OPTION value="80">70 - 80</OPTION>
	  <OPTION value="90">80 - 90</OPTION>
	  <OPTION value="100">90 - 100</OPTION>
	  </SELECT> people attending my party.
	  <INPUT type="submit" value="Submit" name="B1" class="button">
	<P>&nbsp;
	<HR noshade color="#612B83" size="1">
	</FORM>
</TD></tr></TABLE>
<%Else%>

<%
Dim intNumOfPeople, sngPercentage
Dim intMinTonic, intMinGinger, intMinCoke, intMinSoda, intMinOJ, intMinCranberry, intMinLimes, intMinLemons
Dim intMinVodka, intMinWhisky, intMinBourbon, intMinGin, intMinScotch
Dim intTonic, intGinger, intCoke, intSoda, intOJ, intCranberry, intLimes, intLemons
Dim intVodka, intWhisky, intBourbon, intGin, intScotch
Dim strHeader

If Request("number") <> "" Then
	intNumOfPeople = Request("number")
	sngPercentage = Int(intNumOfPeople) / 100
Else
	Response.Redirect ("party_planner.asp")
End If

'   Setup Minimum values
intMinTonic		= 1
intMinGinger	= 1
intMinCoke		= 1
intMinSoda		= 1
intMinOJ		= 1
intMinCranberry = 1
intMinLimes		= 1
intMinLemons	= 1

intMinVodka		= 1
intMinWhisky	= 1
intMinBourbon	= 0
intMinGin		= 1
intMinScotch	= 1

'  Setup actual percentage values
intTonic		= Int(sngPercentage * 12)
intGinger		= Int(sngPercentage * 12)
intCoke			= Int(sngPercentage * 12)
intSoda			= Int(sngPercentage * 8)
intOJ			= Int(sngPercentage * 8)
intCranberry	= Int(sngPercentage * 6)
intLimes		= Int(sngPercentage * 6)
intLemons		= Int(sngPercentage * 6)

intVodka		= Int(sngPercentage * 4)
intWhisky		= Int(sngPercentage * 3)
intBourbon		= Int(sngPercentage * 1)
intGin			= Int(sngPercentage * 3)
intScotch		= Int(sngPercentage * 3)
%>
<P>See in Cocktail : UK:
  <A HREF="glassware.asp">Glassware</a>, <A href="equipment.asp">Bar Equipment</A>, <A href="decoratingDrinks.asp">Decorate Drinks</A><P>For a party of your size (about <%=intNumOfPeople%> people) it is predicted that you will need at least the following:

<H5>Spirits</H5>
<UL>
  <LI><B><%=Max(intVodka, intMinVodka)%></B> bottle(s) of vodka</LI>
  <LI><B><%=Max(intWhisky, intMinWhisky)%></B> bottle(s) of whisky</LI>
  <%If intBourbon > 0 Then%>
	<LI><B><%=Max(intBourbon, intMinBourbon)%></B> bottle(s) of bourbon</LI>
  <%End IF%>
  <LI><B><%=Max(intGin, intMinGin)%></B> bottle(s) of gin</LI>
  <LI><B><%=Max(intScotch, intMinScotch)%></B> bottle(s) of scotch</LI>
  <LI><B>1</B> bottle of sweet vermouth</LI>
  <LI><B>1</B> bottle of dry vermouth</LI>
  <LI><B>2</B> bottle of rum(1 light/1 dark)</LI>
</UL>

<H5>Extra supplies</H5>
<UL>
  <LI><B><%=intNumOfPeople/2%></B> kgs Ice (yes, you'll need that much!)</LI>
  <LI><B><%=Max(intTonic, intMinTonic)%></B> tonic water</LI>
  <LI><B><%=Max(intGinger, intMinGinger)%></B> ginger ale</LI>
  <LI><B><%=Max(intCoke, intMinCoke)%></B> coca-cola</LI>
  <LI><B><%=Max(intSoda, intMinSoda)%></B> soda water</LI>
  <LI><B><%=Max(intOJ, intMinOJ)%></B> orange juice</LI>
  <LI><B><%=Max(intCranberry, intMinCranberry)%></B> cranberry juice</LI>
  <LI><B><%=Max(intLimes, intMinLimes)%></B> whole limes</LI>
  <LI><B><%=Max(intLemons, intMinLemons)%></B> whole lemons</LI>
  <LI>Olives</LI>
  <LI>Maraschino cherries</LI>
  <LI>Napkins</LI>
  <LI>Cocktail sticks</LI>
  <LI>Any extra mixers that you see fit (i.e. pineapple juice)</LI>
  <LI>Any extra fruit that you see fit (i.e. oranges)</LI>
</UL>

<H5>Liqueurs</H5>
<UL>
  <LI>
  You will also need extra Liqueurs. (For example Blue Curacao).<BR>
  You should have a basic collection of these already.<BR>
  These are essential for complete and correct mixing of cocktails.
  </LI>
</UL>

<H5>Glassware</H5>
<UL>
  <LI>
  <B><%=3 * intNumOfPeople%></B> glasses (3 per person)</LI>
  <LI>Stock up on mainly lowball and highball glasses but with a few variations too if possible.<BR>
  <A HREF="glassware.asp">Get more help on selecting glasses</a></LI>
</UL>

<H5>Tools</H5>
<UL>
  <LI>You will  need a basic set of tools for your party.</LI>
  <LI><A HREF="equipment.asp">Click here to view the bar equipment page</A></li>
</UL>

<H5>Preparation</H5>
<UL>
	<LI><P><B>Fruit</B><BR>
	Cut lemons &amp; limes and place in rock glasses or bowls between<BR>
	mixers and spirits. Place extra beverage napkins in front of table <BR>
	for your guest use. If time permits, serve drinks with beverage napkins.<BR>
	Cut extra fruit up for decorative purposes and place in separate bowls.<br>For tips decorating drinks <A HREF="decoratingDrinks.asp">click here</A>.
	</LI>
	<BR>&nbsp;
	<LI><P><B>Spirits</B><BR>
	Take top of bottles and save them. Place in pourers if available.<BR>
	Place bottles so your guests can see what is available.
	</LI>
</UL>
        <p>&nbsp;</td>
      </tr>
</table>

<%End If%><!--#include virtual="/includes/footer.asp" -->