<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:sst000F.asp				_/
'_/	Function	:ステータス配信一覧画面フレーム		_/
'_/	Date			:2003/12/25				_/
'_/	Code By		:aspLand HARA			_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
%>
<!--#include File="Common.inc"-->
<%
'''セッションの有効性をチェック
	CheckLoginH
'''データ取得
	Dim USER,UType
	USER = Session.Contents("userid")
'''エラートラップ開始
	on error resume next
'''DB接続
	Dim ObjConn, ObjRS, StrSQL
	ConnDBH ObjConn, ObjRS

'''ユーザ認証、所属データ所得
	StrSQL = "select UserType,FullName,NameAbrev from mUsers where UserCode='" & USER &"'"
	ObjRS.Open StrSQL, ObjConn
	if err <> 0 then
		DisConnDBH ObjConn, ObjRS	'DB切断
		jumpErrorP "0","c101","01","ステータス配信依頼中一覧","102",""
	end if

	UType = Trim(ObjRS("UserType"))		'ユーザタイプ
	Session.Contents("UType") = Utype
	Session.Contents("LinUN") = Trim(ObjRS("FullName"))		'ログインユーザ名称
	Session.Contents("sUN") = Trim(ObjRS("NameAbrev"))		'ログインユーザ略称

'''DB接続解除
	DisConnDBH ObjConn, ObjRS
'''エラートラップ解除
	on error goto 0
%>
<!-------------ステータス配信一覧画面Frame--------------------------->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>ステータス配信情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<frameset rows="73,*,33" border="0" frameborder="0">
	<noframes>
	このページはフレーム対応のブラウザでご覧ください。
	</noframes>
	<frame src="./sst000T.asp" name="Top" scrolling="no" noresize>
	<frameset cols="120,*" border="0" frameborder="0" border="0" frameborder="0">
		<frame src="./sst000M.asp" name="Menu">
		<frame src="./top.html" name="List" scrolling="no" noresize>
	</frameset>
	<frame src="./sst000B.asp" name="Bottom" scrolling="no" noresize>
</frameset>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------画面終わり--------------------------->
</BODY>
</HTML>
