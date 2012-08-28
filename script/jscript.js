var strTopTitle
function checkSearch(){
	if ( document.search.searchField.value == "" ) {
		alert("Please enter a search query.")
		document.search.searchField.focus()
		return false
	}
	else
		return true
}

function changeTopTitle(strTitle){
	if (document.all&&toptitle){
		strTopTitle	= toptitle.innerHTML
		toptitle.innerHTML = strTitle
	}
}
function changeTopTitleBack(){
	if (document.all&&toptitle){
		toptitle.innerHTML = strTopTitle
	}
}
function setHomePage(){
	if (document.all){
		oHomePage.setHomePage("http://www.cocktail.uk.com/")
	}
	else{
		alert("With your current browser we can not set your home page automatically. Please set manually.")
	}
}
function gotoCategory(){
	location.href = "/shop/" + document.all.shop[shop.selectedIndex].value
}
function checkAddoption(intQuantity, strProduct, frm){
	var selectedItem = frm.selectedIndex;
	var selectedText = frm.options[selectedItem].text;
	var selectedValue = frm.options[selectedItem].value;
	if (selectedValue > 0){
		//if (confirm("Do you want to add "+intQuantity+" x \""+strProduct+" ("+selectedText+")\" to your basket?"))
			return true;
		//else
		//	return false;
	}
	else{
		alert("Please select which type of \""+strProduct+"\" to add\nby clicking the drop down box next to the Buy button.");
		return false;
	}
}
function checkAdd(intQuantity, strProduct){
	//if (confirm("Do you want to add "+intQuantity+" x \""+strProduct+"\" to your basket?"))
		return true;
	//else
	//	return false;
}
function collectionOnly(intQuantity, strProduct)
{

	alert("Sorry, the "+strProduct+" is only available for collection or local delivery, please phone 01223 872769 to check if we can deliver to your area.");
	return false;
}
function displayDeliveryTimes()
{
	delivery = window.open("http://www.drinkstuff.com/member/deliveryTimes.asp","delivery","width=600, height=400, scrollbars=1");
}
function openWin(strArgs){
	strWin = window.open("/db/ingredient_description.asp" + strArgs, "strWin", "height=400, width=600, menubar=0, statusbar=0, scrollbars=1")
	strWin.focus()
}
function addToFavourites(){
	if (document.all){
		window.external.AddFavorite('http://www.cocktail.uk.com','Cocktail : UK')
	}
	else{
		alert("With your current browser we can not add this page to your favourites automatically. Please set manually.")
	}
}
function clearField(){
	var strSearch = document.search.searchField.value
	if (strSearch == "<%=strCocktailSearch%>") {
		document.search.searchField.value = ""
	}
}
