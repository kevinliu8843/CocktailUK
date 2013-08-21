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

'Stock status mods
CONST PRODUCT_STOCK = 1
CONST PRODUCT_SO = 2
CONST PRODUCT_SCR = 3

CONST BASKET_COLUMNS = 20
CONST ITEM_NAME = 0
CONST ITEM_QUANTITY = 1
CONST ITEM_PRICE = 2
CONST TOTAL_PRICE = 3
CONST ITEM_LINK = 4 ' products array - stored the prod ID, basket array, stores the raw prod ID for customised products
CONST ITEM_PRODVERID = 5
CONST VAT_RATE = 6
CONST ITEM_CUSTOM = 7
CONST ITEM_GIFTWRAP = 8
CONST ITEM_OFFERID = 9
CONST ITEM_PRODOFFERID = 10
CONST ITEM_PREORDER = 11
CONST ITEM_STOCK_STATUS = 12
CONST ITEM_PRODID = 13
CONST ITEM_IMAGE = 14

'------------------------------------------------------------------------

strDB    = "Driver={SQL Server}; SERVER=10.0.17.13; Database=cocktailuk_old;UID=cocktailuk;     PWD=pdKe6d#0"
strDBMod = "Driver={SQL Server}; SERVER=10.0.17.13; Database=cocktailuk_old;UID=cocktailukwrite;PWD=dj3#s5c$"

'------------------------------------------------------------------------

bIsAdmin = (Trim(LCase(Session("uname"))) ="leetracey")

'------------------------------------------------------------------------
aryGames(0) = "Board"
aryGames(1) = "Card"
aryGames(2) = "Coin"
aryGames(3) = "Coordination"
aryGames(4) = "Dice"
aryGames(5) = "Endurance"
aryGames(6) = "Luck"
aryGames(7) = "Musical"
aryGames(8) = "Ping-Pong"
aryGames(9) = "TV/Movie"
aryGames(10) = "Speed"
aryGames(11) = "Verbal"
aryGames(12) = "Other"

blnApplyDeliveryZoneRestrictions = False
blnApplyDeliveryCountryRestrictions = False

aryRandomProducts(0) = 1
aryRandomProducts(1) = 2
aryRandomProducts(2) = 75
aryRandomProducts(3) = 421
aryRandomProducts(4) = 59
aryRandomProducts(5) = 80
aryRandomProducts(6) = 475
aryRandomProducts(7) = 149
aryRandomProducts(8) = 72
aryRandomProducts(9) = 187
aryRandomProducts(10) = 382
%>