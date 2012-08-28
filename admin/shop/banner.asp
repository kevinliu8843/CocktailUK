<%option explicit%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/clsUpload.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<%
Dim strMode, strCode, dblCodeValue, strUseOnce, intUseOnce, dblPercentage, intID, intActID
Dim fso, intIDToUse, strType, strFile, objUpload, intImageSize, intImgID
Dim submit_form, strURL, strExtension, strBannerURL, strBannerTarget, strButStyle, strButStyleRed
Dim strDate, cn

intID = Request("ID")
if intID = "" Then
	Response.redirect("default.asp")
End If

If Request("mode") = "" Then
	strMode = "new"
Else
	strMode = Request("mode")
End If
submit_form = Request("submit_form")
Response.CacheControl = "no-cache"
strTitle = "Upload Category Banner"
%>
<!--#include virtual="/includes/header.asp" -->
<SCRIPT language="Javascript">
function submitform(frm)
{
	frm.action  = "banner.asp?ID=<%=intID%>&mode=new&submit_form=true&bannerurl='"+frm.bannerurl.value+"'&bannertarget="+frm.bannertarget.value
	frm.submit()
}
</SCRIPT>

<BODY>

 <H2>Upload Category Banner</H2>
 <table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber3">
   <tr>
     <td width="100%">&nbsp;<%
Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
cn.Open strDBMod
rs.open "SELECT name from dscategory WHERE ID=" & intID, cn, 0, 3
%>
 <P align="center"><B>Category: <A href="/admin/shop/category.asp?mode=update&ID=<%=intID%>&name=<%=rs("name")%>"><%=rs("name")%></A></B></P>
 <%
 rs.close
If Request("submit_form") = "true" Then
	Select case strMode
		case "new"
			'Get file type...
			Set objUpload = New clsUpload
			
			If objUpload.Fields("File1").FileName <> "" Then
				Set fso = Server.CreateObject("Scripting.FileSystemObject")
				strType =  LCase(fso.GetExtensionName(objUpload.Fields("File1").FileName))
				If fso.FileExists(Server.MapPath("/img/banners/"&intID&"."&strType)) Then
					Call fso.DeleteFile(Server.MapPath("/img/banners/"&intID&"."&strType), True)
				End If
				Set fso = nothing
				
				'Update database here...
				strSQL = "UPDATE dscategoryactual SET bannertype='"&strType&"', bannerURL='"&strIntoDB(Replace(request.queryString("bannerurl"), "'", ""))&"', bannertarget='"&strIntoDB(request.queryString("bannertarget"))&"' WHERE catID="&intID
				cn.execute(strSQL)
				
				'Upload image here...
				strFile = objUpload.Fields("File1").FileName
				objUpload("File1").SaveAs Server.MapPath("/images/shop/banners/") & "/" & intID & "." & strType

				Response.write("<P><FONT color=red size=+1>Banner added</FONT>")
			End If
			Set objUpload = Nothing

		case "edit"
			strSQL = "UPDATE dscategoryactual SET bannerURL='"&strIntoDB(request("bannerurl"))&"', bannertarget='"&strIntoDB(request("bannertarget"))&"' WHERE catID="&intID
			cn.execute(strSQL)
			Response.write("<P><FONT color=red size=+1>Banner details edited</FONT>")
			strMode = "new"
			
		case "delete"
			strSQL=  "SELECT bannertype from dscategoryactual WHERE catID="&Request("ID")
			rs.open strSQL, cn, 0, 3
			If NOT rs.EOF Then
				strExtension = rs("bannertype")
			Else
				strExtension = "*"
			End If
			rs.close

			strSQL=  "UPDATE dscategoryactual SET bannertype='', bannerURL='', bannertarget='' WHERE catID="&Request("ID")
			cn.execute(strSQL)
			
			Set fso = Server.CreateObject("Scripting.FileSystemObject")
			If fso.FileExists(Server.MapPath("/img/banners/") & "/" & Request("ID") & "." & strExtension) Then
				fso.DeleteFile(Server.MapPath("/img/banners/") & "/" & Request("ID") & ".*")
			End If
			Set fso = nothing
			
			Response.write("<P><FONT color=red size=+1>Banner deleted</FONT>")
			strMode = "new"
	End Select
End If

If strMode <> "new" AND submit_form <> "true" Then
	strSQL = "SELECT * from dscategoryactual WHERE catID=" & intID 
	'response.write strSQL
	rs.Open strSQL, cn, 0, 3
	If NOT rs.EOF Then
		strBannerURL = strOutDB(rs("bannerURL"))
		strBannerTarget = strOutDB(rs("bannertarget"))
		strURL = "/images/shop/banners/" & rs("catID") & "." & rs("bannertype")
		intActID = rs("catID")
	End If
	rs.close
End If

If strMode = "delete" Then
	strButStyle = strButStyleRed
End If

strSQL = "SELECT * from dscategoryactual WHERE bannerURL<>'' AND bannertype<>'' AND catID=" & intID
rs.Open strSQL, cn, 0, 3
%>
<FORM method="POST" action="banner.asp" language="JavaScript" name="FrontPage_Form1" ID="FrontPage_Form1" <%If strMode = "new" Then%>enctype="multipart/form-data" <%End If%>>
<h4>
<INPUT type="hidden" name="mode" value="<%=strMode%>">
    <INPUT type="hidden" name="submit_form" value="true">
    <INPUT type="hidden" name="ID" value="<%=intID%>">
   <b><%=Capitalise(strMode)%>  banner</b>
</h4>
     <DIV align="center">
     <TABLE border="0" cellpadding="2" cellspacing="0" bordercolor="#111111" id="AutoNumber2" width="100%">
   		<%If strMode = "new" Then%>
       <TR>
         <TD align="left"><b>Banner file:</b></TD>
         <TD><INPUT type="file" name="file1" size="20"></TD>
       </TR>
   		<%End If%>
       <TR>
         <TD align="left"><b>Banner Click Through URL:</b></TD>
         <TD><INPUT type="text" name="bannerurl" size="20" value="<%=strBannerURL%>"></TD>
       </TR>
       <TR>
         <TD align="left"><b>Banner Target Frame:</b></TD>
         <TD><INPUT type="text" name="bannertarget" size="20" value="<%=strBannerTarget%>"></TD>
       </TR>
   		<%If strMode <> "new" Then%>
    	<TR>
         <TD colspan="2">
        <p align="center">
      	<IMG src="<%=strURL%>">
      	<%End if%>
        </p>
      	</TD></TR>
     </TABLE>
   </DIV>
   <P align="center">
   <INPUT type="<%if strMode="new" Then%>button<%Else%>submit<%End If%>" value="<%=Capitalise(strMode)%> banner &gt; &gt;" name="B1" <%=strButStyle%> <%if strMode="new" Then%>onclick="submitform(this.form)"<%End If%> size="20" class="button" ></P>
   </FORM>
 <TABLE border="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="100%" id="AutoNumber1">
   <TR>
     <TH bgcolor="#003366"><font color="#FFFFFF">Banner Image</font></TH>
     <TH bgcolor="#003366"><font color="#FFFFFF">Action</font></TH>
   </TR>
   <%While not rs.EOF%>
   <TR>
     <TD valign="top"><A href="<%=strOutDB(rs("bannerURL"))%>" target="<%=strOutDB(rs("bannertarget"))%>"><IMG SRC="/images/shop/banners/<%=rs("catID")%>.<%=rs("bannertype")%>" border="0"></A></TD>
     <TD align="center" nowrap valign="top"><P align="center"><A href="banner.asp?mode=edit&ID=<%=rs("catID")%>">Edit</A> | <A href="banner.asp?mode=delete&ID=<%=rs("catID")%>">Delete</A></P>
     </TD>
   </TR>
   <%
	  rs.movenext
  wend
  %>
 </TABLE>
     </td>
   </tr>
 </table>

<%
rs.Close
cn.Close
Set cn = Nothing
Set rs = Nothing
%><!--#include virtual="/includes/footer.asp" -->

