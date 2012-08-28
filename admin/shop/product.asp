<%
strTitle = "Shop product manager"
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/admin_functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<H2>Shop Product Manager</H2> 
<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open strDB
If Request("mode") = "update" then
	strSQL = "SELECT * FROM dsproductactual WHERE prodID="&Request("ID")
	rs.open strSQL, cn, 0, 3
	blnNew = True
	If NOT rs.EOF Then
		blnNew = False
		strKeywords = strOutDB(rs("keywords"))
		strTitleA = strOutDB(rs("title"))
		strDescription = strOutDB(rs("description"))
		strSdesc = strOutDB(rs("sdesc"))
		strLdesc = strOutDB(rs("ldesc"))
	End If
	rs.close
	strSQL = "SELECT * FROM dsproduct WHERE status=1 AND ID="&Request("ID")
	rs.open strSQL, cn, 0, 3
	If NOT rs.EOF Then
		strKeywordsDS = strOutDB(rs("keywords"))
		strTitleDS = strOutDB(rs("title"))
		strDescriptionDS = strOutDB(rs("description"))
		strSdescDS = strOutDB(rs("sdesc"))
		strLdescDS = strOutDB(rs("ldesc"))
	End If
	rs.close
%>
<DIV align="center">
  <CENTER> 
  <TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber3">
    <TR>
      <TD>
      <FORM method="POST" action="product.asp">
		<H5 align="center">Product: <%=Request("name")%></H5>
        <TABLE border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber4">
          <TR>
            <TD nowrap><B>Keywords:</B></TD>
            <TD><INPUT type="text" name="keywords" size="56" value="<%=strKeywords%>" maxlength="999"></TD>
          </TR>
          <TR>
            <TD nowrap><B>DS Keywords:</B></TD>
            <TD><FONT size="1"><%=strKeywordsDS%></FONT></TD>
          </TR>
          <TR>
            <TD nowrap><B>Title:</B></TD>
            <TD><INPUT type="text" name="title" size="56" value="<%=strtitlea%>" maxlength="200"></TD>
          </TR>
          <TR>
            <TD nowrap><B>DS Title:</B></TD>
            <TD><FONT size="1"><%=strTitleDS%></FONT></TD>
          </TR>
          <TR>
            <TD nowrap><B>Meta Description</B></TD>
            <TD><INPUT type="text" name="description" size="56" value="<%=strdescription%>" maxlength="200"></TD>
          </TR>
          <TR>
            <TD nowrap><B>DS Description:</B></TD>
            <TD><FONT size="1"><%=strDescriptionDS%></FONT></TD>
          </TR>
          <TR>
            <TD nowrap valign="top"><B>Short Description:</B></TD>
            <TD><TEXTAREA name="sdesc" cols="42" rows="5"><%=strSDesc%></TEXTAREA></TD>
          </TR>
          <TR>
            <TD nowrap valign="top"><B>Long Description:</B></TD>
            <TD><TEXTAREA name="ldesc" cols="42" rows="10"><%=strLdesc%></TEXTAREA></TD>
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
	strSQL = "INSERT into dsproductactual (prodID, keywords, title, description, sdesc, ldesc) VALUES("&Request("ID")&", '"&strIntoDB(Request("keywords"))&"', '"&strIntoDB(Request("title"))&"', '"&strIntoDB(Request("description"))&"', '"&strIntoDB(Request("sdesc"))&"', '"&strIntoDB(Request("ldesc"))&"')"
	cn.execute(strSQL)
Else
	strSQL = "UPDATE dsproductactual SET keywords='"&strIntoDB(Request("keywords"))&"', title='"&strIntoDB(Request("title"))&"', description='"&strIntoDB(Request("description"))&"', sdesc='"&strIntoDB(Request("sdesc"))&"', ldesc='"&strIntoDB(Request("ldesc"))&"', datemodified=GETDATE() WHERE prodID="&Request("ID")
	cn.execute(strSQL)
End If
%>
<P align="center"><FONT color="#FF0000"><I>Product updated</I></FONT></P>
<%End If%>
<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <TR>
    <TD width="50%" valign="top" style="border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1">
    <H5 align="center">Undefined Products</H5>
    <OL>
	<%
	strSQL = "SELECT name, ID from dsproduct where status=1 AND ID NOT IN (SELECT prodID from dsproductactual)"
	rs.open strSQL, cn, 0, 3
	WHILE NOT rs.EOF
	%>
    	<LI><A href="product.asp?mode=update&ID=<%=rs("ID")%>&name=<%=strOutDB(rs("name"))%>"><%=strOutDB(rs("name"))%></A></LI>
    <%	rs.movenext
    wend
    rs.close
    %>
    </OL>
    </TD> 
    <TD width="50%" valign="top">
    <H5 align="center">Defined Products</H5>
    <TABLE border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
      <TR>
        <TH width="50%"><FONT size="1">Name</FONT></TH>
        <TH width="10%"><FONT size="1">Keywords</FONT></TH>
        <TH width="10%"><FONT size="1">Title</FONT></TH>
        <TH width="10%"><FONT size="1">Desc</FONT></TH>
        <TH width="10%"><FONT size="1">SDesc</FONT></TH>
        <TH width="10%"><FONT size="1">LDesc</FONT></TH>
      </TR>
	<%
	strSQL = "SELECT dsproduct.name, dsproduct.ID, dsproductactual.keywords, dsproductactual.title , dsproductactual.description, dsproductactual.sdesc, dsproductactual.ldesc from dsproduct, dsproductactual where dsproductactual.prodID=dsproduct.ID AND status=1 AND dsproductactual.ID IN (SELECT prodID from dsproductactual)"
	rs.open strSQL, cn, 0, 3
	WHILE NOT rs.EOF
	%>
      <TR>
        <TD width="50%"><A href="product.asp?mode=update&ID=<%=rs("ID")%>&name=<%=strOutDB(rs("name"))%>"><%=strOutDB(rs("name"))%></A></TD>
        <TD width="10%" align="center"><%If strOutDB(rs("keywords")) <> "" Then%><img src="/images/shop/gotstock.gif"><%Else%><img src="/images/shop/gotnostock.gif"><%End if%></TD>
        <TD width="10%" align="center"><%If strOutDB(rs("title")) <> "" Then%><img src="/images/shop/gotstock.gif"><%Else%><img src="/images/shop/gotnostock.gif"><%End if%></TD>
        <TD width="10%" align="center"><%If strOutDB(rs("description")) <> "" Then%><img src="/images/shop/gotstock.gif"><%Else%><img src="/images/shop/gotnostock.gif"><%End if%></TD>
        <TD width="10%" align="center"><%If strOutDB(rs("sdesc")) <> "" Then%><img src="/images/shop/gotstock.gif"><%Else%><img src="/images/shop/gotnostock.gif"><%End if%></TD>
        <TD width="10%" align="center"><%If strOutDB(rs("ldesc")) <> "" Then%><img src="/images/shop/gotstock.gif"><%Else%><img src="/images/shop/gotnostock.gif"><%End if%></TD>
      </TR>
    <%	rs.movenext
    wend
    rs.close
    %>
    </TABLE>
    </TD>
  </TR>
</TABLE>
<%
cn.Close
Set cn = Nothing
Set rs = Nothing
%><!--#include virtual="/includes/footer.asp" -->