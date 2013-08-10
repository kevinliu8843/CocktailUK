<!--#include file="errors.asp" -->
<%

If NOT Session("admin") Then
	On Error Resume Next
End If

Dim strSuperScript, rsGlobal, bhideAds, strHideAds, strImage
Dim strTitlePrepend, strTitleAppend, iKounter, strKeywords, user, passwd
Dim intItems, dblValue, blnSkyscraper
Dim strMetaDescription, strMetaKeywords, strMetaTitle

'Do Ads management here====================================================================================
blnSkyscraper = True
bHideAds = False

If InStr(LCase(Request("SCRIPT_NAME")), "google.asp") > 0 Then
	bHideAds = True
End If

If Request("type") = "8" Then
	blnXXX = True
  bHideAds = True
End If

Call DisplayPageLocation(strTitle, strTitleOut, strTopTitle, Request.ServerVariables("SCRIPT_NAME") & "?" & Request.QueryString, "color: white; font-weight: bold; text-decoration: none;")
strToptitle = strTitlePrepend & " " & strToptitle & " " & strTitleAppend

intItems = Max(Session("numberItems")+1,0)
dblValue = Session("valueItems")
If strMetaDescription = "" Then
  strMetaDescription = Replace(strTopTitle," > ",", ") 
End If
If strMetaKeywords = "" Then
  strMetaKeywords = "cocktails, cocktail, "&Replace(strTopTitle," > ",", ")
End If 
If strMetaTitle <> "" Then
  strTopTitle = strMetaTitle
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<link href="/style/style.css" type="text/css" rel="stylesheet">
<meta name="robots" content="ALL">
<meta name="description" content="<%=strMetaDescription%>">
<%If strKeywords = "" Then%>
	<meta name="keywords" content="<%=strMetaKeywords%>">
<%Else%>
	<meta name="keywords" content="<%=strKeywords%>">
<%End If%>
<meta name="revisit-after" content="3 day">
<meta name="distribution" content="GLOBAL">
<meta name="Googlebot" content="all">
<meta name="abstract" content="<%=strTitle%>">
<meta http-equiv="content-language" content="EN">
<meta name="google-site-verification" content="pncNZRLgGxSNLD_-xHUvcx6z6di9D_pU_Kzo-Ldf1kc" />
<meta name="verify-v1" content="j1KzW+k9z2ZccTw61qVc0227g3bZhen6ZCqPR541JsQ=">
<title><%=strTopTitle%></title>
<script type="text/javascript">

  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', 'UA-17242925-1']);
  _gaq.push(['_trackPageview']);

  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();

</script>
</head>

<body>

<div class="wrapper">
  <div class="header">
    <table id="table3" cellspacing="0" cellpadding="0" width="100%" bgcolor="#ffffff" border="0">
     <tr>
      <td align="left" width="85">
      <a href="http://www.cocktail.uk.com/">
      <img border="0" src="../images/cuk_03.jpg" width="85" height="85" alt="Classic cocktails and bar equipment uk"></a></td>
      <td align="left" style="width: 230px">
        <a href="http://www.cocktail.uk.com/">
        <img alt="Classic cocktails and bar equipment uk" src="../images/cuk_07.gif" border="0" width="210" height="32"></a></td>
      <td align="center">
    	  <div id="search_box">
    			<form action="http://www.cocktail.uk.com/sitesearch/google.asp" id="search_form" style="margin: 0px; padding: 0px; display: inline; ">
    			    <input type="hidden" name="cx" value="partner-pub-4852715527905431:j32r2u95lwx">
    			    <input type="hidden" name="cof" value="FORID:10">
    			    <input type="hidden" name="ie" value="UTF-8"> 
    			    <input type="text" name="q" id="SearchField" placeholder="Search for cocktails" class="swap_value"><input type="image" src="../images/template/button_search_go.gif" id="go" name="sa" alt="Search" title="Search">
    			</form>
    		</div>
      </td>
      <td align="right" style="padding-right: 20px; width: 150px;" nowrap>
          <a href="/shop/basket.asp"><img alt="My Basket" src="../images/template/basket_icon.gif" width="36" height="36" align="right" border="0"><strong><u>My Basket</u></strong><br>
          <span style="text-decoration: none; white-space: nowrap;"><%=intItems%> Item<%If intItems <> 1 then%>s<%end if%> 
          &nbsp;&pound;<%=FormatNumber(dblValue,2)%></SPAN></a>
      </td>
     </tr>
    </table>

    <%If NOT bHideAds Then%>
      <div class="topads">
        <div class="ad1">
          <!-- JS AdJug Publisher Code -->    
          <script language="JavaScript">    
          document.write('<scr'+'ipt language="JavaScript" src="http://hosting.adjug.com/AdJugSearch/PageBuilder.aspx?ivi=V3.0+JS&aid=492&slid=49281&height=60&width=468&HTMLOP=False&ShowIFrame=True&CacheBuster=' + Math.floor(Math.random()*99999999) + '"></scr'+'ipt>');
          </script>    
          <noscript>    
          <iframe width="468" height="60" name="AdSpace49281" src="http://hosting.adjug.com/AdJugSearch/PageBuilder.aspx?ivi=V3.0+JS+NS&aid=492&slid=49281&height=60&width=468&HTMLOP=True" frameborder="0" marginwidth="0" marginheight="0" vspace="0" hspace="0" allowtransparency="true" scrolling="no">
          </iframe>    
          </noscript>    
          <!-- JS AdJug Publisher Code -->
        </div>
        <div class="ad2">
          <!--START MERCHANT:merchant name Drinkstuff.com from affiliatewindow.com.-->
          <a href="http://www.awin1.com/cread.php?s=23053&v=8&q=273&r=176043"><img src="http://www.awin1.com/cshow.php?s=23053&v=8&q=273&r=176043" 
          border="0"></a>
          <!--END MERCHANT:merchant name Drinkstuff.com from affiliatewindow.com-->
        </div>
      </div>
    <%End If%>
  </div>

  <div class="leftnav">
    <!--#INCLUDE virtual="/includes/lhs_cocktail.asp"-->
  </div>

  <div class="content">
    <div class="breadcrumb">
      <%If LCase(Request.ServerVariables("SCRIPT_NAME")) = "/default.asp" Then%> 
       Cocktail : UK, cocktails, <span lang="en-gb">cocktail</span> recipes and bar equipment from the UK
      <%Else%>
       <font color="white"><%=strTitleOut%></font>
      <%End If%>
    </div>
    <!--C-->
    <!---->
    <!--/C-->
    <%
    If NOT Session("admin") Then
     Call TrapErrors()
    End If
    %>
  </div>
</div>

<div class="footer" id="footer" align="center">
  <iframe src="//www.facebook.com/plugins/likebox.php?href=http%3A%2F%2Fwww.facebook.com%2Fcocktailuk&amp;width=700&amp;colorscheme=light&amp;show_faces=true&amp;border_color&amp;stream=false&amp;header=true&amp;height=290" scrolling="no" frameborder="0" style="border:none; overflow:hidden; width:625px; height:290px; background-color: white; margin-top: 10px;" allowTransparency="true"></iframe>
</div>

<%If blnSkyscraper AND NOT bHideAds AND NOT blnXXX Then%>
  <div class="skyscraper">
    <!-- Simple IF AdJug Publisher Code -->    
    <iframe width="160" height="600" name="AdSpace49282" src="http://hosting.adjug.com/AdJugSearch/PageBuilder.aspx?ivi=V3.0+IF&aid=492&slid=49282&height=600&width=160&CacheBuster=[time_stamp]&HTMLOP=True" frameborder="0" marginwidth="0" marginheight="0" vspace="0" hspace="0" allowtransparency="true" scrolling="no">
    </iframe>    
    <!-- Simple IF AdJug Publisher Code -->
  </div>
<%End If%>

</body>
</html>