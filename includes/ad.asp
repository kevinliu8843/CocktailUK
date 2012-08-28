<%
Dim blnAdviva, blnAdvertisingUK, blnNormalAd, intAdNumber, intCamp1, intCamp2, intCamp3, intCamp4, intCamp5, intCamp6, intCamp7
Dim blnAccelerator

Randomize
intAdNumber = Int((100 * Rnd) + 1)
' %ages of campains.
intCamp1 = 50 ' AMAZON
intCamp2 = 0  ' DRINKSTUFF
intCamp3 = 0  ' NEEDAPRESENT
intCamp4 = 0  ' POTSTILL
intCamp5 = 50 ' WAILLIAM HILL
intCamp6 = 0 ' CASINO ON NET
intCamp7 = 0 ' FREE IPODS

blnAdviva = False
blnAdvertisingUK = False
blnAccelerator = False

If intAdNumber < intCamp1 Then
	response.write "<A HREF=""http://www.amazon.co.uk/exec/obidos/redirect-home?tag=cocktailheaven&site=amazon""><IMG SRC=""/images/amazon/banner.gif"" HEIGHT=""60"" WIDTH=""468""></A>"
	
ElseIf intAdNumber < intCamp2 + intCamp1 Then
	response.write "<iframe src=""http://www.awin1.com/awshow.php?mid=8&gid=273&aid=10724&iframe=1"" width=468 height=60 frameborder=0 border=0 scrolling=no marginheight=0 marginwidth=0></iframe>"
	
ElseIf intAdNumber < intCamp3 + intCamp2 + intCamp1 Then
	response.write "<A HREF=""http://www.awin1.com/awclick.php?mid=182&gid=1675&id=10724&p=""><IMG src=""/images/needapresent_big.gif""></A>"
	
ElseIf intAdNumber < intCamp1 + intCamp2 + intCamp3 + intCamp4 Then
	response.write "<script language=""JavaScript"" src=""http://www.awin1.com/awshow.php?mid=299&gid=2706&aid=10724""></SCRIPT>"
	
ElseIf intAdNumber < intCamp1 + intCamp2 + intCamp3 + intCamp4  + intCamp5 Then
	response.write "<a href=""http://travis.willhill.com/re.asp?name=WHSD&camp=REF384_0""><img src=""http://www.williamhillcasino.com/banner/uk/468x60_bonus_roul.gif"" width=""468"" height=""60"" border=""0"" alt=""Play now at the William Hill Casino"" /></a>"
	
ElseIf intAdNumber < intCamp1 + intCamp2 + intCamp3 + intCamp4  + intCamp5 + intCamp6 Then
	response.write "<a href=""http://tracker.tradedoubler.com/click?p=14894&a=533010&g=87920"" target=""_blank""><img src=""http://impgb.tradedoubler.com/imp/img/87920/533010"" border=0></a>"
	
ElseIf intAdNumber < intCamp1 + intCamp2 + intCamp3 + intCamp4  + intCamp5 + intCamp6 + intCamp7 Then
	response.write "<a href=""http://www.freeiPods.com/?r=17388134"" target=""_blank""><img src=""/images/freeipods.gif"" border=0></a>"
	
Else
	If blnAdviva Then
%>
	
        <script language="JavaScript">
        <!--
        var ans_timestamp = (new Date()).getTime();
        document.write("<scr" + "ipt language=\"JavaScript\" src=\"http://ads.adviva.net/serve/v=300;m=2;l=1414;ts=" + ans_timestamp + "\"></scr" + "ipt>");
        // -->
        </script>
	
	<%ElseIf blnAccelerator Then%>
	
	<SCRIPT TYPE="text/javascript" LANGUAGE="JavaScript">
	
	// Cache-busting and pageid values
	var random = Math.round(Math.random() * 100000000);
	if (!pageNum) var pageNum = Math.round(Math.random() * 100000000);
	
	document.write('<SCR');
	document.write('IPT TYPE="text/javascript" LANGUAGE="JavaScript" SRC="http://ads.accelerator-media.com/jserver/acc_random=' + random + '/SITE=COCKTAIL.UK/AAMSZ=BANNER/AREA=BUSINESS_FINANCE/pageid=' + pageNum + '">');
	document.write('</SCR');
	document.write('IPT>');
	</SCRIPT>
	
	<%ElseIf blnAdvertisingUK Then%>
	
	<!-- ---------- Advertising.com Banner Code -------------- --> 
	<SCRIPT LANGUAGE=JavaScript> 
	var bnum=new Number(Math.floor(99999999 * Math.random())+1); 
	document.write('<SCR'+'IPT LANGUAGE="JavaScript" '); 
	document.write('SRC="http://servedby.advertising.com/site=137229/size=468060/bnum='+bnum+'/optn=1"></SCR'+'IPT>'); 
	</SCRIPT> 
	<!-- ---------- Copyright 2000, Advertising.com ---------- -->	
	
	<%End If%>
<%End If%>


