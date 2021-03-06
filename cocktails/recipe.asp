﻿<%
Option Explicit

Dim cn, intID, strIngredients, FSO, strName, bFileExists, strType, strDescription
Dim blnUserCocktail, strUserName, i, blnDuplicated, aryReviews, intNumReviews
Dim blnEdit, strDescriptionEDIT, strImgSrc, strImgName, intCase, aryDrink
Dim strCocktailName, intServes, intAccessed, intRate, intUsers, strCat
Dim strMakeCocktail, blnMakeIt, intMaxReviews, objProduct, strURL
%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/product.asp" -->
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
	Call Do301Redirect(strURL)
End If

If Request("delete") <> "" AND session("admin") Then
  cn.execute("UPDATE Cocktail SET status=0 WHERE ID=" & Request("delete"))
  cn.close
  Response.redirect(Request("back"))
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
	Response.Redirect("/cocktails/recipe.asp?ID=1")
End If

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
strTitle = Capitalise(aryDrink(0)) & " " & Capitalise(aryDrink(7))
strMetaDescription = "" & aryDrink(0) & " " & aryDrink(7) & " recipe. Full ingredients & instructions on how to make a " & aryDrink(0) & " " & aryDrink(7) & "."
%>
<!--#include virtual="/includes/header.asp" -->

<h1><%=Capitalise(aryDrink(0)) & " " & Capitalise(aryDrink(7)) & " Recipe"%>
<%If Session("admin") Then%>
  <a target="_top" href="/admin/default.asp?goto=cocktaileditor/default.asp?ID=<%=intID%>" style="font-size: 16px;">Edit</a>
  <a target="_top" href="?delete=<%=intID%>&back=<%=Request.ServerVariables("HTTP_REFERER")%>" style="font-size: 16px;" onclick="return(confirm('Are yo sure you wish to delete this cocktail?'))">Delete</a>
<%End If%>
</h1>

<h4>How to make a <%=LCase(aryDrink(0))%>:</h4>
<p style="margin-bottom: 30px;"><%=aryDrink(1)%></p>

<div class="row collapse">
  <div class="large-3 small-7 column">
    <div style="padding-right: 15px; margin-bottom: 30px;">
      <h5 id="ingredients">Ingredients:</h5>
      <div style="margin-left: 1.5em;"><%=aryDrink(2)%></div>
    </div>
  </div>

  <div class="large-3 small-5 column">
    <div style="padding-right: 15px; margin-bottom: 30px;">
      <h5 id="equipment">Details:</h5>
      <div style="margin-bottom: 5px;">Serves <%=aryDrink(3)%></div>
      <div style="margin-bottom: 5px;">Rated: <%Call displayRatingGraphOnly( CStr(aryDrink(5)) )%></div>
      <div style="margin-bottom: 5px;">Views: <%=aryDrink(4)%></div>
      <%If aryDrink(10) <> "" Then%>
        <div style="margin-bottom: 5px;">Submitter: <%=aryDrink(10)%></div>
      <%End If%>
    </div>
  </div>
  <div class="large-3 small-8 column">
    <div style="padding-right: 15px; margin-bottom: 30px;">
      <%displayRatingPanel%>
    </div>
  </div>
  <div class="large-3 small-4 column">
    <%=aryDrink(11)%>
  </div>
</div>

<div class="row collapse" style="margin-bottom: 40px;">
    <%If blnDuplicated Then%>
        <div class="column large-1"><a href="/account/userHotList.asp?remove=<%=intID%>"><i class="general foundicon-minus" style="font-size: 200%">&nbsp;</i></a></div>
        <div class="column large-11"><a href="/account/userHotList.asp?remove=<%=intID%>">Remove from your favourite recipes</a></div>
    <%Else%>
        <div class="column large-1"><a href="/account/userHotList.asp?add=<%=intID%>"><i class="general foundicon-plus" style="font-size: 200%">&nbsp;</i></a></div>
        <div class="column large-11"><a href="/account/userHotList.asp?add=<%=intID%>">Add to your favourite recipes</a></div>
    <%End If%>
    <%
    If Session("logged") Then 
        Response.write strMakeCocktail
    End If
    %>
</div>

<div class="row collapse">
  <div class="large-9 small-12 column">
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
  </div>
</div>

<div class="row collapse">
  <div class="large-12 column">
    <script type="text/JavaScript">
    AdJug_AID = 492;
    AdJug_SiteAdSpaceID = 49378;
    AdJug_IFrame = false;
    AdJug_ShowDebug = false;
    AdJug_Height = 250;
    AdJug_Width = 300;
    </script>
    <script type="text/JavaScript" src="http://hosting.adjug.com/JavaScript/AdOffer/IncludeResults.js"></script>
  </div>
</div>

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
<form action="/account/addrating.asp" method="post" id="ratingform">
    <input type="hidden" name="ID" value="<%=intID%>">
    <h5>How do you rate it?</h5>
    <div class="row collapse">
        <div class="column small-2"><input type="radio" name="R1" value="5" id="5stars"></div>
        <div class="column small-10"><label for="5stars"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"></label></div>
    </div>
    <div class="row collapse">
        <div class="column small-2"><input type="radio" name="R1" value="4" id="4stars"></div>
        <div class="column small-10"><label for="4stars"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"></label></div>
    </div>
    <div class="row collapse">
        <div class="column small-2"><input type="radio" name="R1" value="3" id="3stars"></div>
        <div class="column small-10"><label for="3stars"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"></label></div>
    </div>
    <div class="row collapse">
        <div class="column small-2"><input type="radio" name="R1" value="2" id="2stars"></div>
        <div class="column small-10"><label for="2stars"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"></label></div>
    </div>
    <div class="row collapse">
        <div class="column small-2"><input type="radio" name="R1" value="1" id="1stars"></div>
        <div class="column small-10"><label for="1stars"><img src="/images/sitesearch/1.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"><img src="/images/sitesearch/0.gif" height="12" width="13" border="0"></label></div>
    </div>
    <div><a href="#" class="button small" onclick="document.getElementById('ratingform').submit()">Submit Rating</a></div>
</form>

 <%If Request("rate") = "true" Then%>
    <div class="alert-box success">Rating Added</div>
 <%elseif Request("rate") = "false" then%>
    <div class="alert-box alert">Please specify a rating</div>
 <%End If%>
<%End Function%>