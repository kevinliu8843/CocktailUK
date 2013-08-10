<%
Option Explicit

Dim cn, intID, strIngredients, FSO, strName, bFileExists, strType, strDescription
Dim blnUserCocktail, strUserName, i, blnDuplicated, aryReviews, intNumReviews
Dim blnEdit, strDescriptionEDIT, strImgSrc, strImgName, intCase, aryDrink
Dim strCocktailName, intServes, intAccessed, intRate, intUsers, strCat, strXXX
Dim strMakeCocktail, blnMakeIt, intMaxReviews, objProduct, strURL
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/shop/functions.asp" -->
<!--#include virtual="/includes/shop/product.asp" -->
<!--#include virtual="/includes/cocktail_functions.asp" --><%
If Request("ID") <> "" AND IsNumeric(Request("ID")) Then
	intID = Int(Request("ID"))
Else
	Response.Redirect("/")
End If

set cn	= Server.CreateObject("ADODB.Connection")
Set rs	= Server.CreateObject("ADODB.RecordSet")
cn.Open strDB

Call PrettyURLRedirectCocktail(cn, rs, intID, strURL)
If strURL <> "" Then
	cn.close
	Set cn = Nothing
	Set rs = Nothing
	Call PermanentRedirect(strURL)
End If

intMaxReviews = Request("reviews")
If intMaxReviews = "" OR NOT IsNumeric(intMaxReviews) Then
	intMaxReviews = 3
Else
	intMaxReviews = Int(Min(intMaxReviews, 10000))
End If

' Get details for this drink ------------------------------------------------------
If NOT GetDrink(rs, cn, intID, aryDrink) Then
	cn.Close
	Set rs	= Nothing
	set cn	= Nothing
	Response.Redirect("/db/viewCocktail.asp?ID=1")
End If

blnXXX = (aryDrink(9) = "XXX rated")

If Session("logged") Then
	blnDuplicated = isInFavourites(intID)
End If

'Get user's reviews...
rs.open "CUK_GETCOCKTAILREVIEWS @id=" & intID, cn, 0, 3
If NOT rs.EOF Then
	aryReviews = rs.GetRows()
	intNumReviews = UBound(aryReviews,2)
Else
	ReDim aryReviews(0,0)
	intNumReviews = -1
End If
rs.close

If Session("logged") Then
	blnMakeIt = canIMakeIt(cn, rs, intID, Session("ID"), strMakeCocktail)
End If

cn.Close
set cn	= Nothing
set rs			= Nothing

blnHardwireTitle = True
strTitle = aryDrink(0) & " " & aryDrink(7) & " recipe. Ingredients and full instructions on how to make it."
strMetaDescription = "" & aryDrink(0) & " " & aryDrink(7) & " recipe. Full ingredients & instructions on how to make a " & aryDrink(0) & " " & aryDrink(7) & "."
%>
<!--#include virtual="/includes/header.asp" -->
<style type="text/css">
 ul { margin-left: 5px; padding-left: 0px; }
 ul { margin-top: 0; }
 li { margin-left: 1em; }
</style>
<h2><%=Capitalise(aryDrink(0) & " " & aryDrink(7)) & " Recipe"%><%If bIsAdmin Then%><a target="_top" class="linksin" href="/admin/default.asp?goto=cocktaileditor/default.asp?ID=<%=intID%>"> 
Edit</a><%End If%></h2>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber6">
  <tr>
    <td valign="top">
    <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber7" height="100%">
      <tr>
        <td width="50%" valign="top">
        <div id="directions">
          <table cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
              <td class="arrowblock" align="left" width="1%" nowrap>
              <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
              <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;INSTRUCTIONS</b></td>
            </tr>
          </table>
          <%=aryDrink(1)%></div>
&nbsp;</td>
        <td width="50%" valign="top">
        <div id="ingredients">
          <table cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
              <td class="arrowblock" align="left" width="1%" nowrap>
              <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
              <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;INGREDIENTS</b></td>
            </tr>
          </table>
          <%=aryDrink(2)%>
        </div>
        </td>
      </tr>
      <tr>
        <td width="100%" colspan="2">
        <table cellspacing="0" cellpadding="0" width="100%" border="0">
          <tr>
            <td class="arrowblock" align="left" width="1%" nowrap>
            <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
            <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;EQUIPMENT 
            NEEDED TO MAKE THIS <%=UCase(aryDrink(7))%></b></td>
          </tr>
        </table>
        <div align="center">
          <img src="/images/pixel.gif" height="5" width="1"><br>
          <table border="0" cellpadding="0" style="border-collapse: collapse" width="100%">
            <tr>
              <td nowrap><%If aryDrink(7)="shooter" Then%><a href="/shop/products/search.asp?search=iceshot"><img border="0" src="/images/drinkstuff/Cocktail%20Equipment/shot_rock.jpg" alt="Shot Rock - Ice shot glasses" width="40" height="40"></a>
              <a href="/shop/products/search.asp?search=shot float kit">
              <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/shot_float.jpg" alt="Shot Float Kit - Help you to layer shooters easier" width="25" height="40"></a>
              <%else%> <a href="/shop/products/search.asp?search=glass">
              <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/glasses.jpg" alt="Glassware" width="28" height="40"></a>
              <a href="/shop/products/search.asp?search=cocktail shaker">
              <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/shaker.jpg" alt="Professional Cocktail Shaker - used in the industry" width="40" height="40"></a>
              <%End if%> <a href="/shop/products/search.asp?search=pourer">
              <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/steel_pourer.jpg" alt="Stainless Steel Pourer - pours ingredients gently onto a drink" width="44" height="40"></a>
              <a href="/shop/products/search.asp?search=measure">
              <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/bar_measures.jpg" alt="Professional Bar Measures - measure out the perfect quantity" width="34" height="40"></a><a onmouseover="show_text('Professional Measures')" onmouseout="hide_text()" href="/shop/products/search.asp?search=measures">
              </a></td>
              <td>
              <div align="right">
                <table border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" id="table2">
                  <tr>
                    <td align="right">
                    <a style="text-decoration: none" class="linksin" href="/shop/basket.asp">
                    View my basket<br>
                    <%=intItems%> item<%If intItems <> 1 then%>s<%end if%> <%=FormatNumber(dblValue,2)%></a></td>
                    <td>
                    <p align="center"><a href="/shop/basket.asp">
                    <img src="/images/shop/view_basket_small.gif" border="0" alt="View my basket"></a></p>
                    </td>
                  </tr>
                </table>
              </div>
              </td>
            </tr>
          </table>
          <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber9">
            <tr>
              <td width="100%">
              <p align="center">
              <img src="/images/pixel.gif" height="5" width="1"><br>
              <a href="/shop/">As a leading bar equipment supplier in the UK, we 
              can deliver to you next working day! <b>Come shopping</b></a><br>
              <img src="/images/pixel.gif" height="20" width="1"><br>
              <img src="/images/pixel.gif" height="5" width="1"></p>
              </td>
            </tr>
          </table>
        </div>
        </td>
      </tr>
      <tr>
        <td width="100%" colspan="2"><%call writeSearchForm%></td>
      </tr>
      <tr>
        <td width="100%" colspan="2">
        <img src="/images/pixel.gif" height="20" width="1"><br>
        <table cellspacing="0" cellpadding="0" width="100%" border="0">
          <tr>
            <td class="arrowblock" align="left" width="1%" nowrap valign="top">
            <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
            <td class="baselightred" width="99%" style="padding-left: 2">
            <b class="contentHeader">YOUR COMMENTS: <%=UCase(aryDrink(0)) & " " & UCase(aryDrink(7))%></b></td>
          </tr>
        </table>
        <img src="/images/pixel.gif" height="5" width="1"><br>
        <%If intNumReviews < 0 Then%> No one has commented on the <%=aryDrink(0)& " " & aryDrink(7)%> yet... <%End If%> <%For i=0 TO Min(intNumReviews,intMaxReviews-1)%>
        <b><%=Capitalise(aryReviews(3,i))%> said:</b><br>
        <img border="0" src="../images/inset_quotebegin.gif"> <%=Trim(aryReviews(4,i))%>
        <img border="0" src="../images/inset_quoteend.gif"><br>
        <%If i<>Min(intNumReviews,intMaxReviews-1) Then%><br>
        <%End If%> <%Next%> <%If intNumReviews>intMaxReviews Then%> <br>
        There are more comments,
        <a href="viewCocktail.asp?ID=<%=intID%>&reviews=999">read them all...</a>
        <%End if%>
        <p align="center"><b>Love or hate this drink?
        <a href="#" onclick="window.open('review.asp?ID=<%=intID%>','review','width=450, height=450, menubar=0, status=0, resizable=1'); return false">
        Have your say...</a></b>
        </p>

<script type="text/JavaScript">
AdJug_AID = 492;
AdJug_SiteAdSpaceID = 49378;
AdJug_IFrame = false;
AdJug_ShowDebug = false;
AdJug_Height = 250;
AdJug_Width = 300;
</script>
<script type="text/JavaScript" src="http://hosting.adjug.com/JavaScript/AdOffer/IncludeResults.js"></script>

        </td>
      </tr>
    </table>
    </td>
    <td width="160" valign="top">
    <table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber8">
      <tr>
        <td width="100%">
        <table cellspacing="0" cellpadding="0" width="100%" border="0">
          <tr>
            <td class="arrowblock" align="left" width="1%" nowrap>
            <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
            <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;PICTURE</b></td>
          </tr>
        </table>
        <img src="/images/pixel.gif" height="5"><br>
        <%=aryDrink(11)%> </td>
      </tr>
      <tr>
        <td width="100%"><%displayRatingPanel%></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td colspan="2">
    <table cellspacing="0" cellpadding="0" width="100%" border="0">
      <tr>
        <td class="arrowblock" align="left" width="1%" nowrap>
        <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
        <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;REGISTERED 
        MEMBERS MENU</b></td>
      </tr>
    </table>
    <table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
      <%If Session("logged") Then%>
      <tr>
        <td><%IF NOT blnDuplicated Then%><a href="/db/member/userHotList.asp?add=<%=intID%>"><img border="0" src="../images/favourites.gif" width="40" height="36"></a><%Else%><a href="/db/member/userHotList.asp?remove=<%=intID%>"><img border="0" src="../images/favourites.gif" width="40" height="36"></a><%End If%></td>
        <td width="100%"><%IF NOT blnDuplicated Then%><a href="/db/member/userHotList.asp?add=<%=intID%>">Add 
        to my favourites</a><%Else%><a href="/db/member/userHotList.asp?remove=<%=intID%>">Remove 
        from my favourites</a><%End If%></td>
      </tr>
      <%End If%>
      <tr>
        <td colspan="2"><%If Session("logged") Then 
								Response.write strMakeCocktail
							Else
								%>
        <table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" id="table3">
          <tr>
            <td><a href="/db/member/userHotList.asp?add=<%=intID%>"><img border="0" src="../images/favourites.gif" width="40" height="36"></a></td>
            <td width="100%"><a href="/db/member/userHotList.asp?add=<%=intID%>">Add to my favourites</a></td>
          </tr>
        </table>
        <p><%End If%></p>
        </td>
      </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td colspan="2">
    <table border="0" cellpadding="0" style="border-collapse: collapse" id="table1">
      <tr>
        <td width="174" valign="top"><a href="/shop/">
        <img border="0" src="../images/homepage/shophomepage_01.jpg" width="174" height="52"></a></td>
        <td valign="top" height="40" style="background:url(../images/homepage/shophomepage_02.gif) top left no-repeat;"><a href="/shop/">
        <img border="0" src="../images/pixel.gif" width="302" height="42"></a></td>
      </tr>
      <tr>
        <td width="174" valign="top"><a href="/shop/">
        <img border="0" src="../images/homepage/shophomepage_03.jpg" alt="Bar accessory shop" width="174" height="122"></a></td>
        <td valign="top">We are the UK&#39;s premier online bar equipment supplier. 
        Ideal for home bar enthusiasts and cocktail connoisseurs alike. We&#39;ve taken 
        over 100,000 orders since 1999 online.<p align="center">
        <!--#include virtual="/includes/shop/categoriesoption.asp" --></p>
        </td>
      </tr>
    </table>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->

<%
Function isInFavourites(intID)
	Dim arySplit
	'Returns whether the cocktail is in the user favourites
	If IsNumeric(Session("ID")) Then
		strSQL = "SELECT count(*) FROM usrfav WHERE cocktailID="&intID&" AND memID=" & Session("ID")
		rs.Open strSQL, cn
		isInFavourites = (rs(0) > 0)
		rs.Close
	End If
End Function

Function displayRatingPanel
%>
<form action="/db/member/addrating.asp" method="post" style="text-align: left">
 <input type="hidden" name="ID" value="<%=intID%>">
 <table cellspacing="0" cellpadding="0" width="100%" border="0">
   <tr>
     <td class="arrowblock" align="left" width="1%" nowrap>
     <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
     <td class="baselightred" width="99%"><b class="contentHeader">&nbsp;DETAILS</b></td>
   </tr>
 </table>
 Serves : <b><%=aryDrink(3)%></b><br>
 Type : <b><%=Capitalise(aryDrink(7))%></b><br>
 Category : <b><%=Capitalise(aryDrink(8))%></b><br>
 Viewed : <b><%=aryDrink(4)%> times</b><br>
 <%If aryDrink(10) <> "" Then%>Submitter : <b><%=aryDrink(10)%></b><br>
 <%End If%> Rated: <%call displayRatingGraphOnly( CStr(aryDrink(5)) )%><br>
 <%If Request("rate") = "true" Then%><font color="#FF0000"><i>Rating Added</i></font>
 <%elseif Request("rate") = "false" then%><font color="#FF0000"><i>Please specify 
 a rating</i></font><%End If%>
 <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
   <tr>
     <td valign="middle" align="center">1</td>
     <td valign="middle" align="center">2</td>
     <td valign="middle" align="center">3</td>
     <td valign="middle" align="center">4</td>
     <td valign="middle" align="center">5</td>
   </tr>
   <tr>
     <td valign="middle" align="center"><input type="radio" name="R1" value="1"></td>
     <td valign="middle" align="center"><input type="radio" name="R1" value="2"></td>
     <td valign="middle" align="center"><input type="radio" name="R1" value="3"></td>
     <td valign="middle" align="center"><input type="radio" name="R1" value="4"></td>
     <td valign="middle" align="center"><input type="radio" name="R1" value="5"></td>
   </tr>
   <tr>
     <td valign="middle" align="center" colspan="5">
     <input type="image" src="../images/main_menus/ratedrink.gif" name="I1" alt="Rate this drink" width="145" height="23" border="0"></td>
   </tr>
 </table>
</form>
<%
End Function

Function sendToFriend
	Dim strName1, strName2, strEmail1, strEmail2
	strName1 = "Your Name"
	strName2 = "Their Name"
	strEmail1 = "Your Email"
	strEmail2 = "Their Email"
	If Request.cookies("CocktailUKSendDrinkName1") <> "" Then
		strName1 = Request.cookies("CocktailUKSendDrinkName1")
	End If
	If Request.cookies("CocktailUKSendDrinkName2") <> "" Then
		strName2 = Request.cookies("CocktailUKSendDrinkName2")
	End If
	If Request.cookies("CocktailUKSendDrinkEmail1") <> "" Then
		strEmail1 = Request.cookies("CocktailUKSendDrinkEmail1")
	End If
	If Request.cookies("CocktailUKSendDrinkEmail2") <> "" Then
		strEmail2 = Request.cookies("CocktailUKSendDrinkEmail2")
	End If
	%>
<form method="POST" action="/mail/mail.asp?type=cocktail">
 <input type="hidden" name="id" value="<%=intID%>">
 <input type="hidden" name="cocktailName" value="<%=capitalise(aryDrink(0))%>">
 <table cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse" bordercolor="#111111">
   <tr>
     <td class="arrowblock" align="left" width="1%" nowrap>
     <img height="16" src="/images/pixel.gif" width="16" border="0"> </td>
     <td class="baselightred" width="93%" nowrap><b class="contentHeader">&nbsp;EMAIL 
     TO A FRIEND</b> </td>
     <td class="baselightred" align="right" width="6%">&nbsp; </td>
   </tr>
 </table>
 <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber5">
   <tr>
     <td>From</td>
     <td>
     <input type="text" name="name1" size="8" value="<%=strName1%>" style="border: 1px solid #612B83; font-size:8pt; color:#612B83" onclick="this.value=''"></td>
     <td>
     <input type="text" name="email1" size="8" value="<%=strEmail1%>" style="border: 1px solid #612B83; font-size:8pt; color:#612B83" onclick="this.value=''"></td>
   </tr>
   <tr>
     <td>To</td>
     <td>
     <input type="text" name="name2" size="8" value="<%=strName2%>" style="border: 1px solid #612B83; font-size:8pt; color:#612B83" onclick="this.value=''"></td>
     <td>
     <input type="text" name="email2" size="8" value="<%=strEmail2%>" style="border: 1px solid #612B83; font-size:8pt; color:#612B83" onclick="this.value=''"></td>
   </tr>
 </table>
 <input type="image" src="../images/main_menus/senddrink.gif" name="s1" alt="Send this drink to your friend" width="145" height="23" border="0">
 <%If Request("mail") <> "" Then%><font color="#FF0000"><center><i>Drink Sent</i></center>
 </font><%End If%> <%
End Function

Function writeSearchForm
%>
</form>
<form action="/sitesearch/default.asp" method="post" name="search2">
 <table cellspacing="0" cellpadding="0" width="100%" border="0"> 
   <tr>
     <td class="arrowblock" align="left" width="1%" nowrap>
     <img height="16" src="/images/pixel.gif" width="16" border="0"></td>
     <td class="baselightred" width="100%"><b class="contentHeader">&nbsp;FIND A 
     DIFFERENT DRINK</b></td>
   </tr>
 </table>
 <div align="center">
   <img src="/images/pixel.gif" height="5" width="1"><br>
   <%IF intID > 1 Then%><a href="?ID=<%=intID-1%>"><%End If%>&laquo; Previous<%IF intID >1 Then%></a><%End If%> |
   <a title="Random drink" href="/db/random.asp">Random</a> |
   <a title="Next drink" href="/db/viewCocktail.asp?ID=<%=intID+1%>">Next &raquo;</a>
   <a href="/db/viewCocktail.asp?ID=<%=intID+1%>"></a><br>
   <img src="/images/pixel.gif" height="5" width="1"><br>
   Search:
   <input type="text" name="searchField" size="24" style="border:1px solid #979797; width: 140px; height: 19px; text-align: left" class="shopoption" value="Drink name here..."><input border="0" src="../images/template/cuk_orange_btn_go.gif" name="I2" width="41" height="19" align="absmiddle" type="image">
 </div>
 </td>
</form>
<%End Function%>