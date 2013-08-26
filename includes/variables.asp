<%
Dim strDB, strMBWDB, strMBWAccounts, bIsAdmin, strIndexShopDB
Dim cnGlobal, strCocktailSearch, strDBMod
Dim intAmazon, blnDisplayHitCounter, blnHardwireTitle
Dim strTitle, strTopTitle , strTitleOut
Dim rs, rs2, strSQL, bPrinterFriendly, aryGames(12), aryRandomProducts(10)
Dim g_aryIngredientType, g_aryIngredientTypeID, g_intNumIngredientTypes
Dim blnApplyDeliveryZoneRestrictions, blnApplyDeliveryCountryRestrictions 

blnHardwireTitle = False

' Categories ------------------------------------------------------------
Const SPIRIT		= 0
Const LIQUOR		= 1
Const JUICE			= 2
Const MIXER			= 3
Const WINE			= 4
Const FRUIT			= 5
Const FLAVOURINGS	= 6
Const SYRUPS		= 7

g_aryIngredientType	= Array("spirits", "liqueurs", "juices", "mixers", "wines/beers", "garnishes", "flavourings", "syrups")
g_aryIngredientTypeID	= Array(1, 2, 3, 4, 5, 6, 7, 8)
g_intNumIngredientTypes	= UBound(g_aryIngredientTypeID)

'------------------------------------------------------------------------

strDB    = "Driver={SQL Server}; SERVER=10.0.17.13; Database=cocktailuk_old;UID=cocktailuk;     PWD=pdKe6d#0"
strDBMod = "Driver={SQL Server}; SERVER=10.0.17.13; Database=cocktailuk_old;UID=cocktailukwrite;PWD=dj3#s5c$"

'------------------------------------------------------------------------

bIsAdmin = (Trim(LCase(Session("uname"))) ="leetracey")

'------------------------------------------------------------------------
%>