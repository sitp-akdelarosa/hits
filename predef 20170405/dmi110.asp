<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi110.asp				_/
'_/	Function	:事前空搬入入力コンテナ選択画面		_/
'_/	Date		:2003/05/28				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:					_/
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
  WriteLogH "b202", "空搬入事前情報入力","00",""	'CW-046
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前登録・空搬入</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT src="./JS/Common.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
window.resizeTo(200,450);
window.focus();

function GoNext(){
  target=document.dmi110F
  Num=LTrim(target.CONnum.value);
  if(Num.length==0){
    alert("コンテナ番号を記入してください");
    target.CONnum.focus();
    return;
  }
  if(!CheckEisu2(target.CONnum.value)){
    alert("コンテナ番号に半角英数字以外の文字を記入しないでください");
    target.CONnum.focus();
    return;
  }
  chengeUpper(target);
  target.submit();
}
//2008-01-31 Add-S M.Marquez
function finit(){
    document.dmi110F.CONnum.focus();
}
//2008-01-31 Add-E M.Marquez
// -->
</SCRIPT>

<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" onload="finit();">
<!-------------空搬入入力コンテナ選択画面--------------------------->
<TABLE border=0 cellPadding=3 cellSpacing=3 width="100%">
  <TR>
    <TD height="300" align=center>
      <FORM name="dmi110F" method="POST" action="./dmi115.asp">
        <B>コンテナ番号で指示</B><BR>
	  <INPUT type=text  name="CONnum" maxlength=12><BR>
	  <A HREF="JavaScript:GoNext()">実行</A><P>
	<A HREF="JavaScript:window.close()">閉じる</A><P>
      </FORM>
  </TD></TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
