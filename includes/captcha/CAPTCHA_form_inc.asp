<%
'****************************************************************************************
'**  Web Wiz Guide - Web Wiz CAPTCHA
'**  
'**  http://www.webwizCAPTCHA.com
'**
'**  Copyright 2005-2006 Bruce Corkhill All Rights Reserved.
'****************************************************************************************            
%>
<script language="javaScript">
function reloadCAPTCHA() {
	document.getElementById('CAPTCHA').src='/includes/CAPTCHA/CAPTCHA_image.asp?'+Date();
}
</script>           
<table width="100%" border="0" cellspacing="0" cellpadding="0">
 <tr>
  <td align="center">
	<img src="/includes/CAPTCHA/CAPTCHA_image.asp" alt="Code Image" id="CAPTCHA" align="absmiddle" />&nbsp;<a href="javascript:reloadCAPTCHA();">Load 
	new code</a></td>
 </tr>
 <tr>
  <td align="center">
	<input type="text" name="securityCode" id="securityCode" size="12" maxlength="12" autocomplete="off" style="width: 90%" /></td>
 </tr></table> 