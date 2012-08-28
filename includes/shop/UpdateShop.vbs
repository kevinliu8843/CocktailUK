Option Explicit

Dim objWinHttp, strURL


Set objWinHttp = CreateObject("Microsoft.XmlHttp")

strURL = "http://www.cocktail.uk.com/includes/sitemap.asp"
objWinHttp.Open "GET", strURL, False
objWinHttp.Send

If objWinHttp.Status <> 200 Then
	Err.Raise 1, "HttpRequester", "Invalid HTTP Response Code - " & objWinHttp.Status 
End If

strURL = "http://www.cocktail.uk.com/shop/update/updateproducts.asp?Get=true&force=true"
objWinHttp.Open "GET", strURL, False
objWinHttp.Send

If objWinHttp.Status <> 200 Then
	Err.Raise 1, "HttpRequester", "Invalid HTTP Response Code - " & objWinHttp.Status 
End If

Set objWinHttp = Nothing

