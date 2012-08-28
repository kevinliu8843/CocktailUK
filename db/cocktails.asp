<%
Option Explicit
strTitle = "Select Cocktail Base"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->

<H2>Cocktails</H2>
<table border="0" cellpadding="5" style="border-collapse: collapse" width="100%" id="table1">
  <tr>
    <td><img border="0" src="../images/redmartini.jpg" width="125" height="244"></td>
    <td>
<TABLE border="0" cellpadding="5" cellspacing="0" id="table2">
  <TR>
    <TD valign="top" align="left" colspan="3">
      <P align="left"><A href="/db/viewAllCocktails.asp?type=1"><B>View
      all cocktails<BR>
      </B></A>Shows a list of all cocktails.
      </TD>
  </TR>
  <TR>
    <TD valign="top" align="left"><A href="/based_cocktails/vodka.asp"><B>Vodka
      based<BR>
      </B></A>Shows a list of cocktails based on vodka.</TD>
    <TD valign="top" align="left">&nbsp;</TD>
    <TD valign="top" align="left">
      <P align="left"><A href="/based_cocktails/gin.asp"><B>Gin based<BR>
      </B></A>Shows a list of cocktails based on gin.</TD>
  </TR>
  <TR>
    <TD valign="top" align="left">
      <P align="left"><A href="/based_cocktails/brandy.asp"><B>Brandy
      based<BR>
      </B></A>Shows a list of cocktails based on brandy.</TD>
    <TD valign="top" align="left">&nbsp;</TD>
    <TD valign="top" align="left"><A href="/based_cocktails/rum.asp"><B>Rum
      based<BR>
      </B></A>Shows a list of cocktails based on rum.</TD>
  </TR>
  <TR>
    <TD valign="top" align="left">
      <P align="left"><A href="/based_cocktails/whisky.asp"><B>Whisky
      based<BR>
      </B></A>Shows a list of cocktails based on whisky.</TD>
    <TD valign="top" align="left">&nbsp;</TD>
    <TD valign="top" align="left">
      <P align="left"><A href="/based_cocktails/tequila.asp"><B>Tequila
      based<BR>
      </B></A>Shows a list of cocktails based on tequila.</TD>
  </TR>
</TABLE>
    </td>
  </tr>
</table>
&nbsp;<p>&nbsp;</p>
<p>&nbsp;</p>
<p>
<%Call DrawSearchCocktailArea()%><!--#include virtual="/includes/footer.asp" --></p>
