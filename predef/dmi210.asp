<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi210.asp				_/
'_/	Function	:事前空搬出入力画面			_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:C-002	2003/08/06	備考欄追加	_/
'_/	Modify		:3th	2003/01/31	3次変更	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
  WriteLogH "b302", "空搬出事前情報入力","01",""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前登録・空バンピック</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function setParam(target){
  //2016.08.22 H.Yoshikawa Upd Start
  //window.resizeTo(300,200);
  window.resizeTo(600,250); //Edited by AK.DELAROSA 2021/01/13
  //2016.08.22 H.Yoshikawa Upd End
  bgset(target);
  window.focus();

}

function GoNext(){
  target=document.dmi210F;
  Num1=LTrim(target.BookNo.value);
  if(Num1.length==0){
    alert("ブッキング番号を記入してください");
    target.BookNo.focus();
    return;
  }
  if(!CheckEisu(target.BookNo.value)){
    alert("ブッキング番号に半角英数字と半角スペース、「-」、「/」以外の文字を記入しないでください");
    target.BookNo.focus();
    return;
  }
  chengeUpper(target);
  //2006/03/06 mod-s h.matsuda
  target.ShoriMode.value="EMoutUpd";
  target.action="./dmi312.asp";
  //target.action="./dmi215.asp"
  //2006/03/06 mod-e h.matsuda
  target.submit();
}

//ブッキング情報への遷移
function GoBookI(){
  target=document.dmi210F;
  Num1=LTrim(target.BookNo.value);
  if(Num1.length==0){
    alert("ブッキング番号を記入してください");
    target.BookNo.focus();
    return;
  }
  if(!CheckEisu(target.BookNo.value)){
    alert("ブッキング番号に半角英数字と半角スペース、「-」、「/」以外の文字を記入しないでください");
    target.BookNo.focus();
    return;
  }

  //2006/03/06 add-s h.matsuda
  target.ShoriMode.value="EMoutInf";
  //2006/03/06 add-e h.matsuda
  BookInfo(target);

}
//2008-01-31 Add-S M.Marquez
function finit(){
    document.dmi210F.BookNo.focus();
}
//2008-01-31 Add-E M.Marquez

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0"  onLoad="setParam(document.dmi210F);finit();">
<!-------------空搬出情報入力画面--------------------------->
<TABLE border=0 cellPadding=3 cellSpacing=1 width="100%">
 <FORM name="dmi210F" method="POST">
  <TR>
    <TD colspan=2>
        <B>空バンピック情報入力</B><BR>
    </TD><TR>
  <TR>
    <TD><DIV class=bgb>*ブッキングＮｏ．</DIV></TD>
    <TD><INPUT type=text name="BookNo" value="<%=Request("BookNo")%>" maxlength=20 size=27></TD></TR>
  <TR>
    <TD height="100" align=center colspan=2 align=center>
<%'Add-s 2006/03/06 h.matsuda %>
	  <INPUT type=hidden name=ShoriMode value="">
<%'Add-e 2006/03/06 h.matsuda %>
       <INPUT type=hidden name=Mord value="0" >
       <INPUT type=button value="ブッキング情報" onClick="GoBookI()"><P>
       <INPUT type=button value="登録" onClick="GoNext()">
       <INPUT type=button value="閉じる" onClick="window.close()">
  </TD></TR>
 </FORM>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
