<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi310.asp				_/
'_/	Function	:事前実搬入番号入力画面		_/
'_/	Date		:2004/01/31				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:3th	2003/01/31	3次変更	_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%><% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
  WriteLogH "b402", "実搬入事前情報入力","00",""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前登録・実搬入</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(200,400);
window.focus();

function GoNext(){
  strA    = new Array("ブッキング番号","コンテナ番号");
  target=document.dmi310F;
  targetA    = new Array();
  targetA[0] = target.BookNo;
  targetA[1] = target.CONnum;
  for(k=0;k<2;k++){
    Num=LTrim(targetA[k].value);
    if(Num.length==0){
      alert(strA[k]+"を記入してください");
      targetA[k].focus();
      return;
    }
    if(k==0){
      if(!CheckEisu(targetA[k].value)){
        alert(strA[k]+"に半角英数字と半角スペース、「-」、「/」以外の文字を記入しないでください");
        targetA[k].focus();
        return;
      }
    }else{
      if(!CheckEisu2(targetA[k].value)){
        alert(strA[k]+"に半角英数字以外の文字を記入しないでください");
        targetA[k].focus();
        return;
      }
    }
  }
  chengeUpper(target);
  target.submit();
}
//2008-01-31 Add-S M.Marquez
function finit(){
    document.dmi310F.BookNo.focus();
}
//2008-01-31 Add-E M.Marquez

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0  onload="finit();">
<!-------------実搬入番号入力画面--------------------------->
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD height="300" align=center>
<%'Mod-s 2006/03/06 h.matsuda%>
<!-----<FORM name="dmi310F" method="POST" action="./dmi315.asp">--->
      <FORM name="dmi310F" method="POST" action="./dmi312.asp">
	  <INPUT type=hidden name="ShoriMode" value="FLin">
<%'Mod-e 2006/03/06 h.matsuda%>
        <B>ブッキング番号</B><BR>
	  <INPUT type=text  name="BookNo" maxlength=20 size=27><BR>
        <B>コンテナ番号</B><BR>
	  <INPUT type=text  name="CONnum" maxlength=12><P>
	  <A HREF="JavaScript:GoNext()">実行</A><P>
	  <A HREF="JavaScript:window.close()">閉じる</A><P>
      </FORM>
  </TD></TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
