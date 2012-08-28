<%
strTitle = "Shop category manager"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/admin_functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<H2>Shop Category Manager</H2> 
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber5">
  <tr>
    <td width="100%"> 
<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open strDBMod
If Request("mode") = "update" then
	strSQL = "SELECT * FROM dscategoryactual WHERE catID="&Request("ID")
	rs.open strSQL, cn, 3, 3
	blnNew = True
	If NOT rs.EOF Then
		blnNew = False
		strExtraKeywords = strOutDB(rs("extrakeywords"))
		strExtraTitle = strOutDB(rs("extratitle"))
		strNoScript = strOutDB(rs("noscript"))
		strBannerURL = strOutDB(rs("bannerURL"))
		strBannerType = strOutDB(rs("bannertype"))
		strBannerTarget = strOutDB(rs("bannertarget"))
	End If
	rs.close
	strSQL = "SELECT extrakeywords, extratitle FROM dscategory WHERE ID="&Request("ID")
	rs.open strSQL, cn, 0, 3
	If NOT rs.EOF Then
		strExtraKeywordsDS = strOutDB(rs("extrakeywords"))
		strExtraTitleDS = strOutDB(rs("extratitle"))
	End If
	rs.close
%>
<DIV align="center">
  <CENTER> 
  <TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber3">
    <TR>
      <TD>
      <FORM method="POST" action="category.asp">
		<H5 align="center">Category: <%=Request("name")%></H5>
        <TABLE border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber4">
          <TR>
            <TD nowrap><B>Extra Keywords:</B></TD>
            <TD><INPUT type="text" name="extrakeywords" size="56" value="<%=strExtraKeywords%>" maxlength="999"></TD>
          </TR>
          <TR>
            <TD nowrap valign="top"><B>DS Keywords:</B></TD>
            <TD valign="top"><font size=1><%=strExtraTitleDS%></font></TD>
          </TR>
          <TR>
            <TD nowrap><B>Extra Title:</B></TD>
            <TD><INPUT type="text" name="extratitle" size="56" value="<%=strExtratitle%>" maxlength="200"></TD>
          </TR>
          <TR>
            <TD nowrap valign="top"><B>DS Title:</B></TD>
            <TD valign="top"><font size=1><%=strExtraKeywordsDS%></font></TD>
          </TR>
          <TR>
            <TD nowrap valign="top"><B>No Script:</B></TD>
            <TD><TEXTAREA name="noscript" cols="42" rows="5"><%=strNoScript%></TEXTAREA></TD>
          </TR>
          <TR>
            <TD nowrap valign="top"><b>Banner:</b></TD>
            <TD><%If strBannerType <> "" Then%><A href="<%=strBannerURL%>" target="<%=strBannerTarget%>"><IMG src="/images/shop/banners/<%=Request("ID")%>.<%=strBannerType%>"></A><BR><%End If%><%If NOT blnNew Then%><a href="banner.asp?ID=<%=Request("ID")%>"><b>Upload Banner &gt; &gt;</b></a><%End If%></TD>
          </TR>
          </TABLE>
        <P align="center"><INPUT type="submit" value="Update &gt; &gt;" name="B1" class="button" ></P>
        <HR noshade color="#612B83" size="1">
        <INPUT type="hidden" name="ID" value="<%=request("ID")%>">
        <INPUT type="hidden" name="updated" value="true">
        <%If blnNew Then%>
        <INPUT type="hidden" name="new" value="true">
        <%End if%>
      </FORM>
      </TD>
    </TR>
  </TABLE>
  </CENTER>
</DIV>
<%End If%>
<%If Request("updated") <> "" Then%>
<%
If Request("new") = "true" Then
	strSQL = "INSERT into dscategoryactual (catID, extrakeywords, extratitle, noscript) VALUES("&Request("ID")&", '"&strIntoDB(Request("extrakeywords"))&"', '"&strIntoDB(Request("extratitle"))&"', '"&strIntoDB(Request("noscript"))&"')"
	cn.execute(strSQL)
Else
	strSQL = "UPDATE dscategoryactual SET extrakeywords='"&strIntoDB(Request("extrakeywords"))&"', extratitle='"&strIntoDB(Request("extratitle"))&"', noscript='"&strIntoDB(Request("noscript"))&"' WHERE catID="&Request("ID")
	cn.execute(strSQL)
End If
%>
<P align="center"><FONT color="#FF0000"><I>Category updated</I></FONT></P>
<%End If%>
<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <TR>
    <TD width="50%" valign="top" style="border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">
    <H5 align="center">Undefined Categories</H5>
    <OL>
	<%
	strSQL = "SELECT name, ID from dscategory where ID NOT IN (SELECT catID from dscategoryactual) order by catorder"
	rs.open strSQL, cn, 0, 3
	WHILE NOT rs.EOF
	%>
    	<LI><A href="category.asp?mode=update&ID=<%=rs("ID")%>&name=<%=strOutDB(rs("name"))%>"><%=strOutDB(rs("name"))%></A></LI>
    <%	rs.movenext
    wend
    rs.close
    %>
    </OL>
    </TD> 
    <TD width="50%" valign="top">
    <H5 align="center">Defined Categories</H5>
    <TABLE border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
      <TR>
        <TH width="60%"><FONT size="1">Name</FONT></TH>
        <TH width="20%"><font size="1">Kywds</font></TH>
        <TH width="20%"><font size="1">Ttl</font></TH>
        <TH width="20%"><font size="1">NScrpt</font></TH>
        <TH width="20%"><font size="1">Bnr</font></TH>
      </TR>
	<%
	strSQL = "SELECT dscategory.name, dscategory.ID, dscategoryactual.extrakeywords, dscategoryactual.extratitle, dscategoryactual.noscript, dscategoryactual.bannertype from dscategory, dscategoryactual where dscategoryactual.catID=dscategory.ID AND dscategory.ID IN (SELECT catID from dscategoryactual) order by catorder"
	rs.open strSQL, cn, 3, 3
	WHILE NOT rs.EOF
	%>
      <TR>
        <TD width="40%"><A href="category.asp?mode=update&ID=<%=rs("ID")%>&name=<%=strOutDB(rs("name"))%>"><%=strOutDB(rs("name"))%></A></TD>
        <TD width="20%" align="center"><%If strOutDB(rs("extrakeywords")) <> "" Then%><img src="/images/shop/gotstock.gif"><%Else%><img src="/images/shop/gotnostock.gif"><%End if%></TD>
        <TD width="20%" align="center"><%If strOutDB(rs("extratitle")) <> "" Then%><img src="/images/shop/gotstock.gif"><%Else%><img src="/images/shop/gotnostock.gif"><%End if%></TD>
        <TD width="20%" align="center"><%If strOutDB(rs("noscript")) <> "" Then%><img src="/images/shop/gotstock.gif"><%Else%><img src="/images/shop/gotnostock.gif"><%End if%></TD>
        <TD width="20%" align="center"><%If strOutDB(rs("bannerType")) <> "" Then%><img src="/images/shop/gotstock.gif"><%Else%><img src="/images/shop/gotnostock.gif"><%End if%></TD>
      </TR>
    <%	rs.movenext
    wend
    rs.close
    %>
    </TABLE>
    </TD>
  </TR>
</TABLE>
    </td>
  </tr>
</table>

<%
cn.Close
Set cn = Nothing
Set rs = Nothing
%><!--#include virtual="/includes/footer.asp" -->
