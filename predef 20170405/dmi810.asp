<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi810.asp				_/
'_/	Function	:事前空搬出CSV入力ファイル設定		_/
'_/	Date		:2003/05/30				_/
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
  WriteLogH "b302", "空搬出事前情報入力","05",""
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前空搬出CSV入力ファイル設定</TITLE>
<SCRIPT language=JavaScript>
<!--
window.focus();
//CW-025 ADD
function SendCSV(){
  if(document.dmi820F.fln.value.length==0){
    alert("ファイルを指定してください。");
    return;
  }else{
    document.dmi820F.submit();
  }
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY>
<!-------------事前空搬出CSV入力ファイル設定--------------------------->
<B>CSVファイル送信</B>
<CENTER>
  <FORM ACTION="./dmi820.asp" NAME="dmi820F" METHOD="POST" ENCTYPE="multipart/form-data">
    <P>送信するファイルを指定してください<BR>
    <input type="file" name="fln" enctype="multipart/form-data" ><BR>
    <INPUT TYPE="HIDDEN" NAME="perm" SIZE="-1" VALUE="forb"></P>
    <P><INPUT TYPE="Button" VALUE="送信" onClick="SendCSV()">
       <INPUT type=button value="閉じる" onClick="window.close()"></P>
  </FORM>
</CENTER>
</BODY></HTML>