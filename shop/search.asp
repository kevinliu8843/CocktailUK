<%
If Request("what") = "drinkstuff" then
	Response.redirect("/shop/products/search.asp?search="&Request("search"))
Else
	Response.redirect("/shop/shop.asp?action=awSearchProducts&mid=&keywords="&Request("search"))
End If
%>