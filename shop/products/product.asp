<% Option Explicit %>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<!--#include virtual="/includes/shop/basketfunctions.asp" -->
<%
Dim objProd, intID, strActKeywords, strDefaultKeywords, strAlsoBought, strComments, strCategory
intID = Request("ID")
If intID <> "" AND IsNumeric(intID) Then
	Set objProd = New CProduct

	If IsNumeric(Request("catID")) AND Request("catID") <> "" Then
		objProd.SetCategory(Request("catID"))
		Call objProd.GetCategoryName()
		strCategory = objProd.m_strCategoryName
		Call objProd.Reset()
	End If
	
	objProd.SetProductID(intID)
	strTopTitle = objProd.DisplayTopTitle
	strTopTitle = objProd.m_strProductName
	If strCategory <> "" Then
		strTopTitle = strTopTitle & " (From " & strCategory & " in the Cocktail : UK Bar Equipment Shop)"
	Else
		strTopTitle = strTopTitle & " - Cocktail : UK Bar Equipment Shop"
	End If
	strTitle = objProd.DisplayTitle
	call objProd.GetKeywords(strDefaultKeywords, strActKeywords)
	strMetaKeywords = objProd.m_strProductName  & ", " & strActKeywords
	strMetaDescription = objProd.m_strMetaDescription
	strAlsoBought = "" 'GetAlsoBought(intID, objProd.m_strProductName)
	%>
	<!--#include virtual="/includes/header.asp" -->
	<!--#include virtual="/includes/shop/header.asp" -->
    <TABLE border="0" cellpadding="0" style="border-collapse: collapse;" bordercolor="#111111" width="100%" id="AutoNumber1" height="40">
      <TR>
        <TD class="shopheadertitle">
    	<H3><%=objProd.m_strProductName%></H3>
        </TD>
      </TR>
    </TABLE>
    
    <!--Display the products-->
    <%objProd.DisplayProduct()%>
    <!--End products-->
    <%=strAlsoBought%>    

	<%If objProd.m_blnProductExists Then%>
		<TABLE border="0" width="100%" cellspacing="0" cellpadding="4">
			<TR>
				<TD bgcolor="#636388" height="1" background="/images/breadcrumbbg.gif"><FONT color="#FFFFFF"><B>Product reviews</B></FONT></TD>
			</TR>
			<tr><td>
				<%
				If Request("reviews") = "all" Then
					objProd.SetNumReviews(99999)
				End If
				objProd.DisplayReviews()
				%>
				<P align="center"><B><A HREF="#" onMouseOver="top.window.status='Review this product'" onMouseOut="top.window.status=''" onClick="window.open('http://www.drinkstuff.com/products/review.asp?ID=<%=intID%>','review','width=400, height=560, menubar=1, status=1')">Review this product...</A></B><br>&nbsp;</P>
			</td></tr>
		</TABLE>
		<TABLE border="0" width="100%" cellspacing="0" cellpadding="4">
			<TR>
				<TD bgcolor="#636388" height="1" background="/images/breadcrumbbg.gif"><FONT color="#FFFFFF"><B>Some comments from our customers</B></FONT></TD>
			</TR>
			<tr><td>
				<%
				objProd.rs.open "SELECT Top 3 * from jointcustomercomments where site=1 and id>rand()*(select count(*) from jointcustomercomments where site=1) and comments<>'' and status=1", objProd.cn
				While NOT objProd.rs.EOF
					strComments = strComments & Replace(objProd.rs("comments"), "drinkstuff", "<a href=""http://www.cocktail.uk.com/"">Cocktail : UK</a>", 1, -1, 1)
					strComments = Replace(strComments, "drinksstuff", "<a href=""http://www.cocktail.uk.com/"">Cocktail : UK</a>", 1, -1, 1)
					strComments = Replace(strComments, "drinksuff", "<a href=""http://www.cocktail.uk.com/"">Cocktail : UK</a>", 1, -1, 1)
					strComments = strComments&"<BR>&nbsp;&nbsp;&nbsp;<B><font size=1>"&objProd.rs("name")
					If strOutDB(objProd.rs("location")) <> "" Then
						strComments = strComments & " - " & objProd.rs("location")
					End If
					strComments = strComments & "</B></FONT><BR/>"
					objProd.rs.movenext
				Wend
				Response.write strComments
				%>
				</td></tr>
		</TABLE>
	<%
	End If
	Set objProd = Nothing
Else
	Response.Redirect("/")
End If%><!--#include virtual="/includes/shop/footer.asp" --><!--#include virtual="/includes/footer.asp" -->