<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi010.asp				_/
'_/	Function	:事前実搬出入力方法選択画面		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:					_/
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
  WriteLogH "b102", "実搬出事前情報入力(共通)","00",""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前登録・実搬出</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(600,400); // Edited by AK.DELAROSA 2021/01/12
window.focus();

function GoNext(n,m){
  if(n==1){
    Num=LTrim(document.dmi010F.CONnum.value);
    if(Num.length==0){
      alert("コンテナ番号を記入してください");
      document.dmi010F.CONnum.focus();
      return;
    }
    if(!CheckEisu2(document.dmi010F.CONnum.value)){
      alert("コンテナ番号に半角英数字以外の文字を記入しないでください");
      document.dmi010F.CONnum.focus();
      return;
    }
    switch(m){
	case 1:
          document.dmi010F.flag.value="1";
	  break;
	case 2:
          document.dmi010F.flag.value="2";
	  break;
        case 3:
          document.dmi010F.flag.value="3";
        break;
      }
  } else {
    Num=LTrim(document.dmi010F.BLnum.value);
    if(Num.length==0){
      alert("ＢＬ番号を記入してください");
      document.dmi010F.BLnum.focus();
      return;
    }
    if(!CheckEisu(document.dmi010F.BLnum.value)){
      alert("ＢＬ番号に半角英数字と半角スペース、「-」、「/」以外の文字を記入しないでください");
      document.dmi010F.BLnum.focus();
      return;
    }
    document.dmi010F.flag.value="4";
  }
  chengeUpper(document.dmi010F);
  document.dmi010F.submit();
}
//2008-01-29 Add-S M.Marquez
function finit(){
    document.dmi010F.CONnum.focus();
}
//2008-01-29 Add-E M.Marquez

// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onload="finit();">
<!-------------実搬出指定画面--------------------------->
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD align=center>
      <FORM name="dmi010F" method="POST" action="./dmi015.asp">
        <B>コンテナ番号で指示</B><BR>
	  <INPUT type=text  name="CONnum" maxlength=12><BR>
	  <A HREF="JavaScript:GoNext(1,1)">指定あり実行</A><BR>
	  <A HREF="JavaScript:GoNext(1,2)">指定なし実行</A><BR>
<% If Session.Contents("UType")<>5 Then %>
	  <A HREF="JavaScript:GoNext(1,3)">一覧から選択実行</A>
<% End If %>
          <P>
        <B>ＢＬ番号で指示</B><BR>
	  <INPUT type=text  name="BLnum" maxlength=20><BR>
	  <A HREF="JavaScript:GoNext(2,0)">実行</A><P>
	<A HREF="JavaScript:window.close()">閉じる</A><P>
        <INPUT type=hidden name="flag">
      </FORM>
  </TD></TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
