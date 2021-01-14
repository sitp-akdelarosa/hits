<%@LANGUAGE = VBScript%>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst100F.asp				_/
'_/	Function	:ステータス配信依頼中一覧画面フレーム		_/
'_/	Date			:2003/12/25				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<LINK REL="stylesheet" TYPE="text/css" HREF="./style.css">
<TITLE>Import Status Delivery Request</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<!-------------ステータス配信依頼中一覧画面Frame--------------------------->
<frameset rows="100,*,35" border="0" frameborder="0" name="100Frame">
	<frame src="./sst100T.asp" name="Top" scrolling="no" noresize>
	<frame src="./sst100L.asp" name="DList">
	<frame src="./sst100B.asp" name="Bottom" scrolling="no" noresize>
	<noframes>
	このページはフレーム対応のブラウザでご覧ください。
	</noframes>
</frameset>
<BODY>
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
