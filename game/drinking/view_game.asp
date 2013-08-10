<%Option Explicit%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<!--#include virtual="/includes/shop/functions.asp" --><%
Dim i, cn, strType, intID

intID = Request("ID")
If intID = "" OR NOT Isnumeric(intID) then
	Response.Redirect("/game/drinking/")
End If
Set cn = Server.CreateObject("ADODB.Connection")
cn.Open strDB
strSQL = "SELECT * from drinkinggame where status=1 AND ID=" & intID
set rs = cn.execute(strSQL)
If NOT rs.EOF Then
	strTitle = rs("title")
%>
<!--#include virtual="/includes/header.asp" -->
<h2><%=strTitle%></h2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber13">
  <tr>
    <td width="100%">
    <p><b><a href="javascript:history.go(-1)">&lt; &lt; Back</a></b></p>
    <h3 align="left"><b>&nbsp;Game directions</b></h3>
    <blockquote>
      <p align="left"><%=strOutDB(rs("directions"))%></p>
    </blockquote>
    <h3 align="left"><b>&nbsp;Sent in by</b></h3>
    <blockquote>
      <p align="left"><%=strOutDB(rs("submitter"))%></p>
    </blockquote>
    <h3>&nbsp;Rated</h3>
    <blockquote>
      <p><%call displayRating( rs("rating"), rs("peoplerated"))%></p>
    </blockquote>
    <div align="center">
      <center> 
      <form action="/account/addrating.asp" method="post">
       <%If Request("rate") = "true" Then%><font color="#FF0000"><i>Rating Added</i></font><%elseif Request("rate") = "false" then%><font color="#FF0000"><i>Please 
       specify a rating</i></font><%End If%>
       <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
         <tr>
           <td valign="middle" align="center" colspan="5"><b>Rate this game</b></td>
         </tr>
         <tr>
           <td valign="middle" align="center">1</td>
           <td valign="middle" align="center">2</td>
           <td valign="middle" align="center">3</td>
           <td valign="middle" align="center">4</td>
           <td valign="middle" align="center">5</td>
         </tr>
         <tr>
           <td valign="middle" align="center">
           <input type="radio" value="1" name="R1"></td>
           <td valign="middle" align="center">
           <input type="radio" name="R1" value="2"> </td>
           <td valign="middle" align="center">
           <input type="radio" name="R1" value="3"> </td>
           <td valign="middle" align="center">
           <input type="radio" name="R1" value="4"> </td>
           <td valign="middle" align="center">
           <input type="radio" name="R1" value="5"> </td>
         </tr>
         <tr>
           <td valign="middle" align="center" colspan="5">
           <input border="0" src="../../images/main_menus/rategame.gif" name="I3" type="image" alt="Rate this game"></td>
         </tr>
       </table>
       <input type="hidden" name="game" value="true">
       <input type="hidden" name="ID" value="<%=intID%>">
      </form>
      </center>
    </div>
    </div>
    <%
cn.close
Set cn = nothing
%>
    <p align="center"><a href="submit_game.asp">
    <img border="1" src="../../images/main_menus/addyourgame.gif" style="border-style: solid; border-color: #800080" width="150" height="23"></a></p>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" --><%
Else
	cn.close
	Set cn = nothing
	Response.Redirect("/game/drinking/")
End If
%>