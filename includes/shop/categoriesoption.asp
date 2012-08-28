<%strScriptName = Request.ServerVariables("SCRIPT_NAME")%>&nbsp;<SELECT name="shop" ID="shop" class="shopoptioncats" onChange="window.location.href='/shop/' + this.options[this.selectedIndex].value"><OPTION value="default.asp">Select a department...</OPTION><OPTION value="products/new-items.asp"<%If InStr(strScriptName,"new-items") > 0 Then %> SELECTED <%End if%>>New Arrivals</OPTION>
<OPTION value="products/london2012.asp"<%If InStr(strScriptName,"london2012") > 0 Then %> SELECTED <%End if%>>Olympics London 2012 Party</OPTION>
<OPTION value="products/best-selling.asp"<%If InStr(strScriptName,"best-selling") > 0 Then %> SELECTED <%End if%>>Best Selling</OPTION>
<OPTION value="products/bar-equipment.asp"<%If InStr(strScriptName,"bar-equipment") > 0 Then %> SELECTED <%End if%>>Bar</OPTION>
<OPTION value="products/cocktail-accessories.asp"<%If InStr(strScriptName,"cocktail-accessories") > 0 Then %> SELECTED <%End if%>>Cocktail</OPTION>
<OPTION value="products/wine-racks-accessories.asp"<%If InStr(strScriptName,"wine-racks-accessories") > 0 Then %> SELECTED <%End if%>>Wine</OPTION>
<OPTION value="products/glassware.asp"<%If InStr(strScriptName,"glassware") > 0 Then %> SELECTED <%End if%>>Glassware</OPTION>
<OPTION value="products/plastic-glasses.asp"<%If InStr(strScriptName,"plastic-glasses") > 0 Then %> SELECTED <%End if%>>Plastic Glasses</OPTION>
<OPTION value="products/catering-equipment.asp"<%If InStr(strScriptName,"catering-equipment") > 0 Then %> SELECTED <%End if%>>Catering</OPTION>
<OPTION value="products/bar-stools.asp"<%If InStr(strScriptName,"bar-stools") > 0 Then %> SELECTED <%End if%>>Bar Stools</OPTION>
<OPTION value="products/party-stuff.asp"<%If InStr(strScriptName,"party-stuff") > 0 Then %> SELECTED <%End if%>>Party</OPTION>
<OPTION value="products/games-room.asp"<%If InStr(strScriptName,"games-room") > 0 Then %> SELECTED <%End if%>>Games Room</OPTION>
<OPTION value="products/gifts-presents.asp"<%If InStr(strScriptName,"gifts-presents") > 0 Then %> SELECTED <%End if%>>Gifts</OPTION>
<OPTION value="products/home-and-garden.asp"<%If InStr(strScriptName,"home-and-garden") > 0 Then %> SELECTED <%End if%>>Home & Garden</OPTION>
<OPTION value="products/inflatable-hot-tubs.asp"<%If InStr(strScriptName,"inflatable-hot-tubs") > 0 Then %> SELECTED <%End if%>>Inflatable Hot Tubs</OPTION>
<OPTION value="products/coming-soon.asp"<%If InStr(strScriptName,"coming-soon") > 0 Then %> SELECTED <%End if%>>Coming Soon</OPTION>
<OPTION value="products/sale.asp"<%If InStr(strScriptName,"sale") > 0 Then %> SELECTED <%End if%>>Sale</OPTION>
<OPTION value="products/clearance-sale.asp"<%If InStr(strScriptName,"clearance-sale") > 0 Then %> SELECTED <%End if%>>Clearance Sale</OPTION>
<OPTION value="products/bundle-deals.asp"<%If InStr(strScriptName,"bundle-deals") > 0 Then %> SELECTED <%End if%>>Bundles</OPTION>
<OPTION value="products/branded-goods.asp"<%If InStr(strScriptName,"branded-goods") > 0 Then %> SELECTED <%End if%>>Shop by Brand</OPTION>
</SELECT>
