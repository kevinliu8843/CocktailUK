<%Option Explicit
strtitle="Games Section Home"%>
<!--#include virtual="/includes/variables.asp" -->
<!--#include virtual="/includes/functions.asp" -->
<!--#include virtual="/includes/header.asp" -->
<SCRIPT language="JavaScript">
function hangMan(){
	document.getElementById('game').src="/game/hangman/default.asp"
}
function jigsaw(){
	document.getElementById('game').src="/game/jigsaw/jigsaw.asp"
}
function blackjack(){
	document.getElementById('game').src="/game/aspbj/bj1.asp"
}
function pubquiz(){
	document.getElementById('game').src='/game/quiz/default.asp'
}
</SCRIPT>
<H2>Games Section!</H2>
<table border="0" cellpadding="5" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber2">
  <tr>
    <td width="100%">Bored? Of course you are, that's why you want to play a quick game! Well, you can either play hangman, a jigsaw game or a rather cool blackjack game...<br>
&nbsp;<TABLE border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
  <TR>
    <TD align="center" valign="top"><a href="javascript:pubquiz()"><b>THE 
	ULTIMATE PUB QUIZ</a></b></TD>
    <TD align="center" valign="top"><B><A href="javascript:hangMan()">COCKTAIL HANGMAN</A></B></TD>
    <TD align="center" valign="top"><B><A href="javascript:jigsaw()">COCKTAIL JIGSAW</A></B></TD>
    <TD align="center" valign="top"><B><A href="javascript:blackjack()">BLACKJACK</A></B></TD>
  </TR>
  </TABLE>
  <P align="center">
  <IFRAME src="start.asp" width="98%" height="510" ID="game" name="game" name="I1" style="border: 1px solid #FFFFFF" border="0" frameborder="0"></IFRAME>
</P>
    </td>
  </tr>
</table>
<!--#include virtual="/includes/footer.asp" -->