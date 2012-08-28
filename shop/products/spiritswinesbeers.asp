<% Option Explicit %>
<%Dim intID%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<%strtitle="Spirits/Wines/Beers"%>
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" -->
<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <TR>
    <TD valign="top" class="shopheadertitle">
	<H3>Spirits/Beers/Wines</H3>
    </TD>
  </TR>
</TABLE>
<FIELDSET class="fieldsetshop">
  <table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber2" width="100%">
    <tr>
      <td valign="top" width="44"><a href="?category=spirits">
      <img border="0" src="../../images/shop/spirits.gif" align="left" width="44" height="100"></a></td>
      <td valign="top" width="33%"><b><font size="4">
      <a href="?category=spirits">Spirits</a></font></b></td>
      <td valign="top" width="41"><a href="?category=beers">
      <img border="0" src="../../images/shop/beers.gif" align="left" width="41" height="100"></a></td>
      <td valign="top" width="33%"><b><font size="4"><a href="?category=beers">
      Beers</a></font></b></td>
      <td valign="top" width="40"><a href="?category=wine">
      <img border="0" src="../../images/shop/wines.gif" align="left" width="41" height="100"></a></td>
      <td valign="top" width="33%"><b><font size="4"><a href="?category=wine">Wines</a></font></b>
      <SPAN class="linksin"><br><font size="2" class="linksin">&nbsp;- <a class="linksin" href="?category=white wine">white</a>
<BR>&nbsp;- <a class="linksin" href="?category=red wine">red</a><br>
&nbsp;- <a class="linksin" href="?category=rose wine">rose</a><br>
&nbsp;- <a class="linksin" href="?category=sparkling wine">sparkling</a><br>
&nbsp;- <a class="linksin" href="?category=dessert wine">dessert</a><br>
&nbsp;- <a class="linksin" href="?category=other wine">other</a></font></SPAN></td>
    </tr>
    <tr>
      <td valign="top" colspan="6">
  <form method="GET" action="spiritswinesbeers.asp">
   <P align="center"><B><br>
   Search for</B>
            <INPUT style="border:1px solid #612B83; width: 112px; color: #FF0000; height: 21px; text-align: left" maxlength="90" size="10" name="search" onClick="this.value=''" value="<%=Request("search")%>"> 
   In 
   <select size="1" name="in" class="shopoption" >
   <option selected value="spirits">Spirits</option>
   <option value="beers">Beers</option>
   <option value="wine">Wines</option>
   <option value="White wine">Wines - white</option>
   <option value="Red wine">Wines - red</option>
   <option value="Rose wine">Wines - rosé</option>
   <option value="Sparkling wine">Wines - sparkling</option>
   <option value="dessert wine">Wines - dessert wine</option>
   <option value="other wines">Wines - other</option>
   </select>  
            <INPUT type="submit" value="GO" name="searchbutton" style="border:3px double #FF9966; color: #FFFFFF; font-weight: bold; font-size: 9px; background-color: #612B83; height: 21px; font-family:Tahoma"></P>
            <input type="hidden" name="submit_form" value="true">
            </form>
      </td>
    </tr>
    <tr>
      <td valign="top" colspan="6">
  <%
  If Request("category") <> "" Then
SELECT Case LCase(Request("category"))
Case "spirits"
intID = 92
Case "beers"
intID = 90
Case "wine"
intID = 94
Case "white wine"
intID = 95
Case "red wine"
intID = 97
Case "rose wine"
intID = 98
Case "sparkling wine"
intID = 99
Case "dessert wine"
intID = 96
Case "other wine"
intID = 100
End Select
  	%>
  	<H3><%=Capitalise(Request("category"))%></H3> 
  	<script language="JavaScript" src="http://pf.tradedoubler.com/pf/pf?a=533010&categoryId=<%=intID%>&xslUrl=http://www.cocktail.uk.com/includes/shop/shopxsl.xsl&maxResults=100&firstResult=0&js=true"></script>
  	<%
  End If
  If Request("search") <> "" Then
SELECT Case LCase(Request("in"))
Case "spirits"
intID = 92
Case "beers"
intID = 90
Case "wine"
intID = 94
Case "white wine"
intID = 95
Case "red wine"
intID = 97
Case "rose wine"
intID = 98
Case "sparkling wine"
intID = 99
Case "dessert wine"
intID = 96
Case "other wine"
intID = 100
End Select
    	%>
  	<H3>Search for <%=Request("search")%> in <%=Request("in")%></H3>
  	<script language="JavaScript" src="http://pf.tradedoubler.com/pf/pf?a=533010&productName=<%=Request("search")%>&categoryId=<%=intID%>&xslUrl=http://www.cocktail.uk.com/includes/shop/shopxsl.xsl&maxResults=100&firstResult=0&js=true" charset="UTF-8"></script>
  	<%
  End If
  %>
</td>
    </tr>
  </table>
</FIELDSET><!--#include virtual="/includes/shop/footer.asp" --><!--#include virtual="/includes/footer.asp" -->