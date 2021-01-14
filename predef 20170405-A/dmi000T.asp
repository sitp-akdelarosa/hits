<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi000T.asp				_/
'_/	Function	:事前情報一覧画面ヘッダ			_/
'_/	Date		:2003/05/26				_/
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
  dim DayTime,day
  'サーバ日時の取得
  getDayTime DayTime
  day = DayTime(0) & "年" & DayTime(1) & "月" & DayTime(2) & "日" &_
        DayTime(3) & "時" & DayTime(4) & "分現在の情報"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>事前情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<SCRIPT language=JavaScript>
<!--
function OpenCodeWin()
{
  var CodeWin;
  CodeWin = window.open("../codelist.asp?user=<%=Session.Contents("userid")%>","codelist","scrollbars=yes,resizable=yes,width=300,height=350");
  CodeWin.focus();
}
// -->
</SCRIPT>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------事前入力初期画面Top--------------------------->
<TABLE border="0" cellPadding="0" cellSpacing="0" width="100%" height="40">
   <TR>
     <TD rowspan="3" width="506" valign=top><IMG src="Image/predef_title.gif" width="506" height="40"></TD>
     <TD align="right" bgColor="#000099" height="14" colspan="2"><IMG src="Image/logo_hits_ver2.gif" height="14" width="300"></TD>
   </TR>
   <TR>
   	  <TD height="2" colspan="2"></TD>
   </TR>
   <TR>
   	   <TD align="center" valign="top"><%=day%></TD>
       <TD align=right>
	   	<Form>
          <INPUT type="button" value="コード一覧" OnClick="OpenCodeWin()" style="height:22;">
          <SPAN style="width:20;"></SPAN>
        </Form>
	   </TD>
   </TR>
</TABLE>
<!-------------画面終わり--------------------------->
</BODY></HTML>
