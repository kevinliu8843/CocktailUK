<table border="0" cellpadding="0" cellspacing="0" style="border-bottom: 1px solid #626288;" bordercolor="#612B83" id="AutoNumber12" width="100%">
  <tr>
    <td height="35" background="/images/main_menus/theshop.gif" valign="bottom">
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber13" height="100%">
      <tr>
        <td width="100%">
        <a href="/shop/">
        <img border="0" src="../../images/pixel.gif" width="258" height="34" alt="Bar equipment shop"></a></td>
        <td>
        <div align="right">
          <table border="0" style="border-collapse: collapse; margin-right: 5px;" bordercolor="#111111" id="table2" width="150">
            <tr>
              <td align="right" nowrap>
              <a style="text-decoration: none" class="linksin" href="/shop/basket.asp">
              <b><u>View my basket</u></b> (<%=intItems%> Item<%If intItems <> 1 then%>s<%end if%> &pound;<%=FormatNumber(dblValue,2)%>)</a>
              </td>
            </tr>
          </table>
        </div>
        </td>
      </tr>
    </table>
    </td>
  </tr>
  <form method="GET" action="/search/">
   <tr>
     <td height="20" background="/images/grad_write_purple_small.gif">
     <div align="center">
       <center>
       <table border="0" cellpadding="6" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="table1">
           <tr>
           <td nowrap colspan="2">
           <table style="width: 100%" cellspacing="0" cellpadding="0">
			<tr>
				<td>
           <strong>FREE </strong>standard UK <a href="/shop/delivery.asp">delivery</a> on orders over &pound;75</td>
				<td style="text-align: right;">250,000 orders since 1999</td>
			</tr>
			</table>
			</td>
         </tr>
         <tr>
		 <td nowrap><b>Go </b><span><b>to:</b>
           <!--#include virtual="/includes/shop/categoriesoption.asp" --></span>
           </td>
           <td nowrap align="right">
			<div id="search_box">
				<form id="search_form" method="post" action="/search/" style="margin: 0px; padding: 0px; display: inline;">
					<INPUT type="hidden" name="update" value="1">
					<INPUT type="hidden" name="pg" value="1">
					<INPUT type="hidden" name="o" value="10">
					<INPUT type="hidden" name="theshop" value="ON">
			        <input type="text" name="s" id="SearchField" value="Search" class="swap_value" style="width: 150px;" onfocus="this.value=''; document.getElementById('search_box').style.backgroundImage='url(http://www.cocktail.uk.com/images/template/bg_search_box_over.gif)'" onblur="document.getElementById('search_box').style.backgroundImage='url(http://www.cocktail.uk.com/images/template/bg_search_box.gif)'"><input type="image" src="../../images/template/button_search_go.gif" id="go" alt="Search" title="Search">
			    </form>
			</div>
           </td>
         </tr>
       </table>
       </center>
     </div>
     </td>
   </tr>
   </form>
</table>