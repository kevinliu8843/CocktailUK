﻿<%
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
	intMaxReviews = 10
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

blnDuplicated = False
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
strTitle = Capitalise(aryDrink(0)) & " " & Capitalise(aryDrink(7)) & " Recipe"
strMetaDescription = "" & aryDrink(0) & " " & aryDrink(7) & " recipe. Full ingredients & instructions on how to make a " & aryDrink(0) & " " & aryDrink(7) & "."
%>
<!--#include virtual="/includes/header.asp" -->

<style type="text/css">
 ul { margin-left: 5px; padding-left: 0px; }
 ul { margin-top: 0; }
 li { margin-left: 1em; }
</style>

<h1><%=Capitalise(aryDrink(0)) & " " & Capitalise(aryDrink(7)) & " Recipe"%>
<%If Session("admin") Then%>
  <a target="_top" class="linksin" href="/admin/default.asp?goto=cocktaileditor/default.asp?ID=<%=intID%>">Edit</a>
<%End If%>
</h1>

<h3>How to make a <%=LCase(aryDrink(0))%>:</h3>
<div style="margin-bottom: 30px;"><%=aryDrink(1)%></div>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse;" bordercolor="#111111" width="100%" id="AutoNumber7">
  <tr>
    <td width="25%" valign="top">
      <div style="padding-right: 15px; margin-bottom: 30px;">
        <h3 id="ingredients">Ingredients:</h3>
        <div>Serves <%=aryDrink(3)%></div>
        <div><%=aryDrink(2)%></div>
      </div>
    </td>
    <td width="25%" valign="top">
      <div style="padding-right: 15px; margin-bottom: 30px;">
        <h3 id="equipment">You'll also need:</h3>
        <div>
          <%If aryDrink(7)="shooter" Then%>
            <a href="/shop/products/search.asp?search=iceshot"><img border="0" src="/images/drinkstuff/Cocktail%20Equipment/shot_rock.jpg" alt="Shot Rock - Ice shot glasses" width="40" height="40"></a>
            <a href="/shop/products/search.asp?search=shot float kit">
            <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/shot_float.jpg" alt="Shot Float Kit - Help you to layer shooters easier" width="25" height="40"></a>
          <%else%>
            <a href="/shop/products/search.asp?search=glass">
            <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/glasses.jpg" alt="Glassware" width="28" height="40"></a>
            <a href="/shop/products/search.asp?search=cocktail shaker">
            <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/shaker.jpg" alt="Professional Cocktail Shaker - used in the industry" width="40" height="40"></a>
          <%End if%>
          <a href="/shop/products/search.asp?search=pourer">
          <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/steel_pourer.jpg" alt="Stainless Steel Pourer - pours ingredients gently onto a drink" width="44" height="40"></a>
          <a href="/shop/products/search.asp?search=measure">
          <img border="0" src="/images/drinkstuff/Cocktail%20Equipment/bar_measures.jpg" alt="Professional Bar Measures - measure out the perfect quantity" width="34" height="40"></a><a onmouseover="show_text('Professional Measures')" onmouseout="hide_text()" href="/shop/products/search.asp?search=measures">
          </a>
        </div>
      </div>
    </td>
    <td width="25%" valign="top">
      <div style="padding-right: 15px; margin-bottom: 30px;">
        <%displayRatingPanel%>
      </div>
    </td>
    <td width="25%">
      <%=aryDrink(11)%>
    </td>
  </tr>
  <tr>
    <td colspan="4">
      <table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td colspan="2">
            <table border="0" cellpadding="2" cellspacing="0" width="100%" style="border-collapse: collapse" bordercolor="#111111" id="table3">
              <%IF blnDuplicated Then%>
                <tr>
                  <td>
                      <a href="/account/userHotList.asp?remove=<%=intID%>"><img border="0" src="../images/favourites.gif" width="40" height="36"></a>
                  </td>
                  <td width="100%">
                      <a href="/account/userHotList.asp?remove=<%=intID%>">Remove from your favourites</a>
                  </td>
                </tr>
              <%Else%>
                <tr>
                  <td><a href="/account/userHotList.asp?add=<%=intID%>"><img border="0" src="../images/favourites.gif" width="40" height="36"></a></td>
                  <td width="100%"><a href="/account/userHotList.asp?add=<%=intID%>">Add to my favourites</a></td>
                </tr>
              <%End If%>
            </table>
            <%
            If Session("logged") Then 
              Response.write strMakeCocktail
            End If
            %>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="3" style="padding-top: 30px;">
      <div id="disqus_thread"></div>
      <script type="text/javascript">
          /* * * CONFIGURATION VARIABLES: EDIT BEFORE PASTING INTO YOUR WEBPAGE * * */
          var disqus_shortname = 'cocktailuk'; // required: replace example with your forum shortname

          /* * * DON'T EDIT BELOW THIS LINE * * */
          (function() {
              var dsq = document.createElement('script'); dsq.type = 'text/javascript'; dsq.async = true;
              dsq.src = '//' + disqus_shortname + '.disqus.com/embed.js';
              (document.getElementsByTagName('head')[0] || document.getElementsByTagName('body')[0]).appendChild(dsq);
          })();
      </script>
      <noscript>Please enable JavaScript to view the <a href="http://disqus.com/?ref_noscript">comments powered by Disqus.</a></noscript>
      <a href="http://disqus.com" class="dsq-brlink">comments powered by <span class="logo-disqus">Disqus</span></a>
      
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
<form action="/account/addrating.asp" method="post" style="text-align: left">
 <input type="hidden" name="ID" value="<%=intID%>">
 <H3>Details:</h3>
 Type: <%=Capitalise(aryDrink(7))%><br>
 Category: <%=Capitalise(aryDrink(8))%><br>
 Viewed: <%=aryDrink(4)%> times<br>
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
%>