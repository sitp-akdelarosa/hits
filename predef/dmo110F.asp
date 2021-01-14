<%
	@LANGUAGE = VBScript
	@CODEPAGE = 932
%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmo110F.asp				_/
'_/	Function	:空搬入情報一覧画面フレーム		_/
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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>空搬入情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<!-------------空搬入情報一覧画面Frame--------------------------->
<frameset rows="100,*,35" border="0" frameborder="0" name="110Frame">
  <frame src="./dmo110T.asp" name="Top" scrolling="no" noresize>
  <frame src="./dmo110L.asp" name="DList" scrolling="no">
  <frame src="./dmo110B.asp" name="Bottom" scrolling="no" noresize>
  <noframes>
  このページはフレーム対応のブラウザでご覧ください。
  </noframes>
</frameset>
<BODY>
<!-------------画面終わり--------------------------->
</BODY></HTML>
