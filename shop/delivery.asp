<% Option Explicit %>
<%
strTitle = "Cocktail : UK Shop - Delivery Charges, Options & Timescales"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" -->
<style type="text/css">
.nicebgandborder TH
{
	background-image: url('http://www.drinkstuff.com/img/newtemplate/bluegradbg.gif');
	background-repeat: repeat-x;
	background-position: left top;
	border: 1px solid #1F4B94;
	border-collapse: collapse;
	COLOR: #FFFFFF;
	HEIGHT: 25px;
}
h4{
	border-bottom: 1px solid #636388;
}
</style>
<DIV style="padding: 10px;"><%=GetURL("http://www.drinkstuff.com/delivery.asp?contentonly=true&groupID=" & Request("groupID"))%></div>
<!--#include virtual="/includes/footer.asp" -->

<%Response.End%>
