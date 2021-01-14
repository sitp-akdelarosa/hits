<% @LANGUAGE = VBScript %>
<%
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/	SystemName	:Hits					_/
'_/	FileName	:dmi000F.asp				_/
'_/	Function	:事前情報一覧画面フレーム		_/
'_/	Date		:2003/05/26				_/
'_/	Code By		:SEIKO Electric.Co 大重			_/
'_/	Modify		:					_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
%>
<% Option Explicit %>
<%
	'HTTPコンテンツタイプ設定
	Response.ContentType = "text/html; charset=Shift_JIS"
	Response.AddHeader "Pragma","No-Cache"
	Response.AddHeader "Cache-Control","No-Cache"
%>
<!--#include File="Common.inc"-->
<%
'セッションの有効性をチェック
  CheckLoginH
'データ取得
  dim USER,UType
  USER       = Session.Contents("userid")
'エラートラップ開始
    on error resume next
'DB接続
  dim ObjConn, ObjRS, StrSQL
  ConnDBH ObjConn, ObjRS


'ユーザ認証、所属データ所得
  StrSQL = "select UserType,FullName,HeadCompanyCode,NameAbrev from mUsers " &_ 
           "where UserCode='" & USER &"'"
  ObjRS.Open StrSQL, ObjConn
  if err <> 0 then
    DisConnDBH ObjConn, ObjRS	'DB切断
    jampErrerP "0","b101","01","共通：ユーザデータ取得","102",""
  end if

    UType                      = Trim(ObjRS("UserType"))		'ユーザタイプ
    Session.Contents("UType")  = Utype
    Session.Contents("LinUN")  = Trim(ObjRS("FullName"))		'ログインユーザ名称
    Session.Contents("sUN")    = Trim(ObjRS("NameAbrev"))		'ログインユーザ略称
    If UType=5 Then
      Session.Contents("COMPcd") = Trim(ObjRS("HeadCompanyCode"))	'ヘッド会社コード
    Else 
      '2010/04/16 Upd-S Tanaka &nbsp;の文字列がSQL分に使われるので修正
      Session.Contents("COMPcd") = "&nbsp;　"
      'Session.Contents("COMPcd") = ""
      '2010/04/16 Upd-S Tanaka
    End If
'DB接続解除
  DisConnDBH ObjConn, ObjRS
'エラートラップ解除
  on error goto 0
%>
<!-------------事前情報一覧画面Frame--------------------------->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE>事前情報一覧</TITLE>
<META content="text/html; charset=Shift_JIS" http-equiv=Content-Type>
<META content="MSHTML 5.00.2919.6307" name=GENERATOR></HEAD>
<frameset rows="40,*,33" border="0" frameborder="0">
  <noframes>
  このページはフレーム対応のブラウザでご覧ください。
  </noframes>
  <frame src="./dmi000T.asp" name="Top" scrolling="no" noresize>
  <frameset cols="100,*" border="0" frameborder="0" border="0" frameborder="0">
    <frame src="./dmi000M.asp" name="Menu">
<!--
    <frame src="./dmo010F.asp" name="List" scrolling="no" noresize>
-->
    <frame src="./top.asp" name="List" scrolling="no" noresize>
  </frameset>
  <frame src="./dmi000B.asp" name="Bottom" scrolling="no" noresize>
</frameset>
<BODY leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<!-------------画面終わり--------------------------->
</BODY></HTML>
