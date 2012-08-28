<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<%
Dim objProd, intCategory, strMainTitle, strNoScript, strBannerTargetCat, strBannerURLCat, strBannerImageSrcCat
intCategory = 564
If Request("pagesize") <> "" Then
	Session("pagesize") = Request("pagesize")
End If
Set objProd = New CProduct
objProd.SetCategory(intCategory)
strTitleAppend = objProd.DisplayTopTitle
strTitle = objProd.DisplayTitle
strMainTitle = "<H3>" & objProd.DisplayTitle & "</H3>"
strMetaKeywords = objProd.GetCategoryKeywords
strNoScript = objProd.GetNoScript
Call objProd.GetBanner(strBannerTargetCat, strBannerURLCat, strBannerImageSrcCat)
%>
<!--#include virtual="/includes/header.asp" -->
<!--#include virtual="/includes/shop/header.asp" -->
<TABLE border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <TR>
    <TD class="shopheadertitle" valign="top">
	<%=strMainTitle%>
    </TD>
    <TD align="right" class="shoplinebgonly" nowrap valign="top"><FORM style="display: inline; padding-top: 5px; margin: 0px" method="POST" action="<%=Request.ServerVariables("SCRIPT_NAME")%><%If Request.QueryString <> "" Then%>?<%=Request.QueryString%><%End If%>">Show <SELECT size="1" name="pagesize" style="font-size: 8pt; padding-top: 2px;" onChange="this.form.submit()">
      <OPTION value="5" <%If Session("pagesize") = "5" Then%>selected<%End If%>>5</OPTION>
      <OPTION value="10" <%If Session("pagesize") = "10" OR Session("pagesize") = "" Then%>selected<%End If%>>10</OPTION>
      <OPTION value="20" <%If Session("pagesize") = "20" Then%>selected<%End If%>>20</OPTION>
      <OPTION value="999" <%If Session("pagesize") = "999" Then%>selected<%End If%>>All</OPTION>
      </SELECT> products</FORM>
    </TD>
  </TR>
</TABLE>
<%If strBannerURLCat <> "" AND strBannerImageSrcCat <> "" Then%>
	<P align="center"><A href="<%=strBannerURLCat%>" target="<%=strBannerTargetCat%>"><IMG src="<%=strBannerImageSrcCat%>" border="0" align="center"></A></P>
<%End If%>
	<!--Display the products-->
	<%objProd.DisplayProducts%>  
	<!--End products-->
	<%Set objProd = Nothing%>
<!--#include virtual="/includes/shop/footer.asp" -->
<%If strNoScript <> "" Then%>
   <noscript><%If InStr(1, strNoScript, "<H1>", 1) <= 0 Then%><H1><%End If%><%=strNoScript%><%If InStr(1, strNoScript, "<H1>", 1) <= 0 Then%></H1><%End If%></noscript>
<%End If%><!--#include virtual="/includes/footer.asp" -->
